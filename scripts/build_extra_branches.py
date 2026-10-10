#!/usr/bin/env python3
"""Ubah export Excel peserta (sheet "Sudah Bayar") menjadi data/extra-branches.js.

Dipakai untuk cabang yang belum ada di Google Sheet dashboard publik
(Road to FK, Open House 1, Open House 2). Hanya angka agregat per hari yang
ditulis ke file hasil; tidak ada nama, email, atau data peserta lain.

Pakai:
    pip install openpyxl
    python3 scripts/build_extra_branches.py path/ke/export.xlsx

Target, deadline, dan status "ditutup" tiap cabang diatur di BRANCHES di bawah.
"""
import collections
import datetime as dt
import json
import os
import sys

import openpyxl

# (nilai kolom "Cabang" di Excel, nama di dashboard, target, deadline, sudah ditutup)
# Target/deadline None = cabang arsip: tampil sebagai hasil akhir, tanpa progres.
BRANCHES = [
    ("OPEN HOUSE 2", "Open House 2", 450, "2026-11-03", False),
    ("OPEN HOUSE 1", "Open House 1", None, None, True),
    ("ROAD TO FK", "Road to FK", None, None, True),
]

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "extra-branches.js")


def parse_date(value):
    try:
        return dt.datetime.strptime(str(value).strip(), "%d-%m-%Y").date()
    except ValueError:
        return None


def read_rows(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = []
    try:
        for row in wb.active.iter_rows(values_only=True):
            rows.append(row)
    except TypeError:
        # Export ini punya atribut margin halaman kosong di akhir sheet yang membuat
        # openpyxl berhenti; semua baris data sudah terbaca sebelum itu.
        pass
    return rows


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    src = sys.argv[1]
    rows = read_rows(src)
    col = {name: i for i, name in enumerate(rows[0])}
    data = [r for r in rows[1:] if r[0] is not None]

    def when(r):
        # Tren di sheet publik dihitung dari tanggal bayar; kalau kosong pakai tanggal daftar.
        return parse_date(r[col["Tanggal Bayar"]]) or parse_date(r[col["Dibuat Pada"]])

    snapshot = max(d for d in (when(r) for r in data) if d)

    out = []
    for cabang, name, target, deadline, closed in BRANCHES:
        days = collections.Counter(
            when(r) for r in data if str(r[col["Cabang"]]).strip().upper() == cabang
        )
        days.pop(None, None)
        if not days:
            sys.exit(f"Tidak ada baris untuk cabang {cabang!r} di {src}")
        out.append({
            "name": name,
            "target": target,
            "deadline": deadline,
            "closed": closed,
            "total": sum(days.values()),
            "first": min(days).isoformat(),
            "last": max(days).isoformat(),
            "daily": {d.isoformat(): n for d, n in sorted(days.items())},
        })

    # Jumlah pendaftar Olimpiade per region. Dipakai dashboard wilayah untuk memecah baris
    # "Region lain" di sheet (mis. Surabaya dan Online) yang di sheet hanya muncul sebagai satu angka.
    regions = collections.Counter(
        str(r[col["Region"]]).strip().upper()
        for r in data if str(r[col["Cabang"]]).strip().upper().startswith("OLIMPIADE")
    )

    payload = {
        "snapshot": snapshot.isoformat(),
        "source": os.path.basename(src),
        "branches": out,
        "olimpiadeRegions": dict(sorted(regions.items())),
    }
    with open(OUT, "w", encoding="utf-8") as f:
        f.write("// DIBUAT OTOMATIS oleh scripts/build_extra_branches.py. Jangan diedit tangan.\n")
        f.write("window.MEDSPIN_EXTRA = ")
        f.write(json.dumps(payload, ensure_ascii=False, separators=(",", ":")))
        f.write(";\n")

    print(f"Data per {snapshot}")
    for b in out:
        print(f"  {b['name']}: {b['total']} peserta ({b['first']} s.d. {b['last']})")
    print(f"  Olimpiade per region: {len(regions)} region, {sum(regions.values())} peserta")


if __name__ == "__main__":
    main()
