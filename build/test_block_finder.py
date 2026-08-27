"""Checks the block finder used by addin/modWrangleBlocks.bas.

VBA cannot be run outside Excel, so this mirrors FindBlocks in Python and
exercises it against the sheet shapes that turn up in a data room: monthly
blocks down a tab, tables side by side, ragged edges, stray labels.

It verifies the ALGORITHM, not the VBA implementation - the two have to be
kept in step by hand. Re-run after changing either:

    python build/test_block_finder.py
"""

def find_blocks(grid, min_blank_rows=1, min_rows=2, min_cols=1):
    """grid[r][c] is True where the cell has content. 0-based; returns
    (top, bottom, left, right) inclusive, 0-based."""
    n_rows = len(grid)
    n_cols = len(grid[0]) if n_rows else 0
    if not n_rows or not n_cols:
        return []

    row_has = [any(grid[r]) for r in range(n_rows)]

    # 1. split into row bands on runs of blank rows
    bands = []
    r = 0
    while r < n_rows:
        if not row_has[r]:
            r += 1
            continue
        start = last = r
        blank_run = 0
        while r < n_rows:
            if row_has[r]:
                blank_run = 0
                last = r
            else:
                blank_run += 1
                if blank_run >= min_blank_rows:
                    break
            r += 1
        bands.append((start, last))
        r = last + 1 + min_blank_rows        # step past the gap we stopped on

    # 2. within each band, split into column runs on any blank column
    blocks = []
    for top, bottom in bands:
        col_has = [any(grid[rr][c] for rr in range(top, bottom + 1)) for c in range(n_cols)]
        c = 0
        while c < n_cols:
            if not col_has[c]:
                c += 1
                continue
            left = c
            while c < n_cols and col_has[c]:
                c += 1
            right = c - 1
            if (bottom - top + 1) >= min_rows and (right - left + 1) >= min_cols:
                blocks.append((top, bottom, left, right))
    return blocks


def make(rows):
    """'.' blank, anything else content."""
    return [[ch != '.' for ch in row] for row in rows]

def show(bs):
    return [f"r{t+1}-{b+1} c{l+1}-{rr+1}" for t, b, l, rr in bs]

fails = []
def check(name, got, want):
    ok = got == want
    print(f"{'ok  ' if ok else 'FAIL'} {name:44s} {show(got)}")
    if not ok:
        print(f"       wanted {show(want)}")
        fails.append(name)

# two tables one under the other, one blank row between
check("stacked vertically", find_blocks(make([
    "XXX",
    "XXX",
    "...",
    "XXX",
    "XXX",
])), [(0,1,0,2), (3,4,0,2)])

# side by side, one blank column between
check("side by side", find_blocks(make([
    "XX.XX",
    "XX.XX",
])), [(0,1,0,1), (0,1,3,4)])

# a 2x2 arrangement
check("four blocks in a grid", find_blocks(make([
    "XX.XX",
    "XX.XX",
    ".....",
    "XX.XX",
    "XX.XX",
])), [(0,1,0,1), (0,1,3,4), (3,4,0,1), (3,4,3,4)])

# an internal blank line should not split the table when 2 blanks are required
check("internal blank row, minBlank=2", find_blocks(make([
    "XXX",
    "...",
    "XXX",
]), min_blank_rows=2), [(0,2,0,2)])

# ... but does when only 1 is required. min_rows has to come down too, or the
# one-row halves are filtered out - which is the right default, and the reason
# a header-plus-one-row table is the smallest thing reported.
check("internal blank row, minBlank=1", find_blocks(make([
    "XXX",
    "...",
    "XXX",
]), min_blank_rows=1, min_rows=1), [(0,0,0,2), (2,2,0,2)])

check("one-row blocks dropped by default", find_blocks(make([
    "XXX",
    "...",
    "XXX",
]), min_blank_rows=1), [])

check("one table only", find_blocks(make([
    "XXXX",
    "XXXX",
])), [(0,1,0,3)])

check("nothing at all", find_blocks(make(["....", "...."])), [])

# a stray one-row label is dropped by min_rows
check("stray label dropped", find_blocks(make([
    "X..",
    "...",
    "XXX",
    "XXX",
])), [(2,3,0,2)])

# ragged block: a short column still belongs to it
check("ragged right edge", find_blocks(make([
    "XXX",
    "XX.",
    "XXX",
])), [(0,2,0,2)])

# three monthly blocks, the usual data-room shape
check("three monthly blocks", find_blocks(make([
    "XXXX",
    "XXXX",
    "XXXX",
    "....",
    "XXXX",
    "XXXX",
    "XXXX",
    "....",
    "XXXX",
    "XXXX",
    "XXXX",
])), [(0,2,0,3), (4,6,0,3), (8,10,0,3)])

# blocks of different widths stacked
check("different widths stacked", find_blocks(make([
    "XXXXX",
    "XXXXX",
    ".....",
    "XXX..",
    "XXX..",
])), [(0,1,0,4), (3,4,0,2)])

# a gap wider than the minimum
check("two blank rows between", find_blocks(make([
    "XXX",
    "XXX",
    "...",
    "...",
    "XXX",
    "XXX",
])), [(0,1,0,2), (4,5,0,2)])

print()
print("ALL PASSED" if not fails else f"FAILURES: {fails}")
