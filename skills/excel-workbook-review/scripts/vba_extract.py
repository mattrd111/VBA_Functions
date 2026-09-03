#!/usr/bin/env python3
"""Extract VBA source from a workbook or a vbaProject.bin. Pure stdlib.

Usage:
  python3 vba_extract.py <Book.xlsm|vbaProject.bin> [-o OUTDIR] [--json]

Excel keeps macros in xl/vbaProject.bin - an OLE compound file holding
RLE-compressed module streams (MS-CFB + MS-OVBA). This reads both formats
directly so the source can be diffed, linted and reviewed without Excel.

Standard modules land in OUTDIR/*.bas, class modules in *.cls, and document
modules (ThisWorkbook, Sheet1, ...) in OUTDIR/document-modules/*.cls - those
cannot be re-imported as new components, their code has to be pasted back.
"""
import argparse
import json
import struct
import sys
import zipfile
from pathlib import Path

FREESECT, ENDOFCHAIN, FATSECT, DIFSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFC
CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class CompoundFile:
    """Minimal MS-CFB reader: enough to pull named streams out by path."""

    def __init__(self, data):
        if not data.startswith(CFB_SIGNATURE):
            raise ValueError("not an OLE compound file (bad signature)")
        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
        first_dir = struct.unpack_from("<I", data, 0x30)[0]
        first_mini_fat = struct.unpack_from("<I", data, 0x3C)[0]
        first_difat = struct.unpack_from("<I", data, 0x44)[0]
        num_difat = struct.unpack_from("<I", data, 0x48)[0]

        self.fat = self._read_fat(first_difat, num_difat)
        self.mini_fat = self._read_chain_as_uint32(first_mini_fat)
        self.entries = self._read_directory(first_dir)
        root = self.entries[0]
        self.mini_stream = self._read_chain(root["start"], root["size"], mini=False)
        self.paths = {}
        self._walk(root["child"], "")

    # -- sector plumbing ---------------------------------------------------
    def _sector(self, index):
        start = (index + 1) * self.sector_size
        return self.data[start:start + self.sector_size]

    def _read_fat(self, first_difat, num_difat):
        difat = list(struct.unpack_from("<109I", self.data, 0x4C))
        sector = first_difat
        for _ in range(num_difat):
            if sector in (ENDOFCHAIN, FREESECT):
                break
            raw = self._sector(sector)
            values = struct.unpack("<%dI" % (self.sector_size // 4), raw)
            difat.extend(values[:-1])
            sector = values[-1]
        fat = []
        for sec in difat:
            if sec in (FREESECT, ENDOFCHAIN):
                continue
            raw = self._sector(sec)
            fat.extend(struct.unpack("<%dI" % (self.sector_size // 4), raw))
        return fat

    def _chain(self, start, fat):
        out, sector, guard = [], start, 0
        limit = len(fat) + 1
        while sector not in (ENDOFCHAIN, FREESECT) and guard < limit:
            out.append(sector)
            sector = fat[sector] if sector < len(fat) else ENDOFCHAIN
            guard += 1
        return out

    def _read_chain_as_uint32(self, start):
        raw = b"".join(self._sector(s) for s in self._chain(start, self.fat))
        return list(struct.unpack("<%dI" % (len(raw) // 4), raw)) if raw else []

    def _read_chain(self, start, size, mini):
        if mini:
            chunks = [
                self.mini_stream[s * self.mini_sector_size:(s + 1) * self.mini_sector_size]
                for s in self._chain(start, self.mini_fat)
            ]
        else:
            chunks = [self._sector(s) for s in self._chain(start, self.fat)]
        return b"".join(chunks)[:size]

    # -- directory ---------------------------------------------------------
    def _read_directory(self, first_dir):
        raw = b"".join(self._sector(s) for s in self._chain(first_dir, self.fat))
        entries = []
        for off in range(0, len(raw), 128):
            chunk = raw[off:off + 128]
            if len(chunk) < 128:
                break
            name_len = struct.unpack_from("<H", chunk, 0x40)[0]
            name = chunk[: max(0, name_len - 2)].decode("utf-16-le", "replace")
            entries.append({
                "name": name,
                "type": chunk[0x42],
                "left": struct.unpack_from("<I", chunk, 0x44)[0],
                "right": struct.unpack_from("<I", chunk, 0x48)[0],
                "child": struct.unpack_from("<I", chunk, 0x4C)[0],
                "start": struct.unpack_from("<I", chunk, 0x74)[0],
                "size": struct.unpack_from("<Q", chunk, 0x78)[0],
            })
        return entries

    def _walk(self, index, prefix):
        # Directory siblings form a red-black tree; an iterative walk keeps a
        # corrupt or cyclic tree from blowing the stack.
        stack, seen = [index], set()
        while stack:
            i = stack.pop()
            if i in (FREESECT, ENDOFCHAIN) or i >= len(self.entries) or i in seen:
                continue
            seen.add(i)
            entry = self.entries[i]
            path = prefix + entry["name"]
            if entry["type"] == 2:
                self.paths[path] = entry
            elif entry["type"] == 1:
                self._walk(entry["child"], path + "/")
            stack.extend([entry["left"], entry["right"]])

    def read(self, path):
        entry = self.paths.get(path)
        if entry is None:
            for known, value in self.paths.items():
                if known.lower() == path.lower():
                    entry = value
                    break
        if entry is None:
            raise KeyError(path)
        mini = entry["size"] < self.mini_cutoff
        return self._read_chain(entry["start"], entry["size"], mini=mini)


def decompress(data, start=0):
    """MS-OVBA CompressedContainer -> plain bytes."""
    if data[start] != 0x01:
        raise ValueError("bad compressed container signature")
    out = bytearray()
    pos = start + 1
    while pos + 1 < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        pos += 2
        size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        end = pos + size - 2
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < end and pos < len(data):
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end or pos >= len(data):
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                    continue
                token = struct.unpack_from("<H", data, pos)[0]
                pos += 2
                difference = len(out) - chunk_start
                bit_count = 4
                while (1 << bit_count) < difference:
                    bit_count += 1
                bit_count = max(bit_count, 4)
                length_mask = 0xFFFF >> bit_count
                length = (token & length_mask) + 3
                offset = ((token & ~length_mask & 0xFFFF) >> (16 - bit_count)) + 1
                source = len(out) - offset
                if source < 0:
                    raise ValueError("copy token points before the chunk start")
                for i in range(length):
                    out.append(out[source + i])
        pos = end
    return bytes(out)


def parse_dir_stream(raw):
    """Pull module name / stream / text offset / type out of the decompressed dir stream.

    The stream is a record list, but several records lie about their size
    (PROJECTVERSION and the MODULE terminator most notoriously), so walking it
    strictly from byte zero desynchronises on most real files. Instead: find the
    PROJECTMODULES marker, then anchor on each MODULENAME record and parse only
    forward from there. Resynchronising per module means one odd record cannot
    corrupt the rest of the project.
    """
    codepage = 1252
    cp_marker = raw.find(struct.pack("<HI", 0x0003, 2))
    if cp_marker >= 0:
        codepage = struct.unpack_from("<H", raw, cp_marker + 6)[0]

    modules_at = raw.find(struct.pack("<HI", 0x000F, 2))
    if modules_at < 0:
        return [], codepage
    count = struct.unpack_from("<H", raw, modules_at + 6)[0]

    starts, pos = [], modules_at + 8
    while len(starts) < count and pos < len(raw) - 6:
        idx = raw.find(b"\x19\x00", pos)
        if idx < 0:
            break
        size = struct.unpack_from("<I", raw, idx + 2)[0]
        body = raw[idx + 6:idx + 6 + size]
        if 0 < size <= 64 and len(body) == size and all(32 <= b < 127 for b in body):
            starts.append(idx)
            pos = idx + 6 + size
        else:
            pos = idx + 2

    modules = []
    for i, start in enumerate(starts):
        limit = starts[i + 1] if i + 1 < len(starts) else len(raw)
        module = {"name": None, "stream": None, "offset": 0, "type": None}
        pos = start
        while pos + 6 <= limit:
            record_id, size = struct.unpack_from("<HI", raw, pos)
            body = raw[pos + 6:pos + 6 + size]
            if record_id == 0x0019:
                module["name"] = body
            elif record_id == 0x001A:
                module["stream"] = body
            elif record_id == 0x0031 and size >= 4:
                module["offset"] = struct.unpack_from("<I", body, 0)[0]
            elif record_id in (0x0021, 0x0022):
                module["type"] = "standard" if record_id == 0x0021 else "class"
            elif record_id == 0x002B:
                break
            pos += 6 + size
        if module["name"]:
            modules.append(module)
    return modules, codepage


def load_container(path):
    path = Path(path)
    blob = path.read_bytes()
    if blob.startswith(CFB_SIGNATURE):
        return blob
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as z:
            for name in ("xl/vbaProject.bin", "word/vbaProject.bin", "xl/vbaProject.bin.bak"):
                if name in z.namelist():
                    return z.read(name)
        raise SystemExit(f"{path} contains no VBA project (no xl/vbaProject.bin) - the workbook has no macros.")
    raise SystemExit(f"{path} is neither a workbook nor a vbaProject.bin")


DOCUMENT_HINTS = ("VB_Base = \"0{00020819", "VB_Base = \"0{00020820", "VB_Base = \"0{00020P")


def extract(path, outdir=None):
    cfb = CompoundFile(load_container(path))
    modules, codepage = parse_dir_stream(decompress(cfb.read("VBA/dir")))
    encoding = "cp%d" % codepage
    try:
        "x".encode(encoding)
    except LookupError:
        encoding = "cp1252"

    results = []
    for module in modules:
        stream_name = (module["stream"] or module["name"]).decode(encoding, "replace")
        name = module["name"].decode(encoding, "replace")
        try:
            raw = cfb.read("VBA/" + stream_name)
            source = decompress(raw, module["offset"]).decode(encoding, "replace")
        except (KeyError, ValueError, IndexError) as exc:
            results.append({"name": name, "error": str(exc), "kind": "unreadable"})
            continue

        source = source.replace("\r\n", "\n").replace("\r", "\n")
        is_document = any(hint in source for hint in DOCUMENT_HINTS)
        if module["type"] == "standard":
            kind, ext = "standard", ".bas"
        elif is_document:
            kind, ext = "document", ".cls"
        else:
            kind, ext = "class", ".cls"

        code_lines = [
            l for l in source.splitlines()
            if l.strip() and not l.startswith("Attribute ") and not l.strip().startswith("'")
        ]
        results.append({
            "name": name, "stream": stream_name, "kind": kind, "extension": ext,
            "lines": len(source.splitlines()), "code_lines": len(code_lines),
            "source": source,
        })

    if outdir:
        outdir = Path(outdir)
        (outdir / "document-modules").mkdir(parents=True, exist_ok=True)
        for r in results:
            if "source" not in r:
                continue
            folder = outdir / "document-modules" if r["kind"] == "document" else outdir
            target = folder / (r["name"] + r["extension"])
            body = r["source"]
            if not body.startswith("Attribute VB_Name"):
                body = 'Attribute VB_Name = "%s"\n%s' % (r["name"], body)
            target.write_text(body, encoding="utf-8")
            r["written_to"] = str(target)
    return results


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("path")
    ap.add_argument("-o", "--outdir", help="write modules here as .bas/.cls files")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    try:
        results = extract(args.path, args.outdir)
    except (ValueError, KeyError, struct.error) as exc:
        # A damaged, encrypted or unusually-built project is not worth guessing
        # at - say so and point at the reliable route out of Excel itself.
        print(f"Could not read the VBA project: {exc}", file=sys.stderr)
        print("Export the modules from Excel instead (assets/vba_io.bas -> ExportModules), "
              "then review those files.", file=sys.stderr)
        return 3
    if args.json:
        print(json.dumps([{k: v for k, v in r.items() if k != "source"} for r in results], indent=2))
    else:
        for r in results:
            if "error" in r:
                print(f"{r['name']}: UNREADABLE ({r['error']})")
            else:
                print(f"{r['name']:<24} {r['kind']:<10} {r['code_lines']:>5} code lines"
                      + (f" -> {r['written_to']}" if "written_to" in r else ""))
        if not results:
            print("No VBA modules found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
