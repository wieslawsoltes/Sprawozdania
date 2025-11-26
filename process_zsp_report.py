#!/usr/bin/env python3
"""
Analiza zespołów szkolno-przedszkolnych:
- wylicza dzieci na podstawie szkół i przedszkoli w zespole,
- dodaje kolumny wejściowe i formułę w raporcie.
"""

from pathlib import Path

import pandas as pd


def build_address(row: pd.Series) -> str:
    """Składa adres z kolumn ulicy, numeru i poczty."""
    street = str(row["Ulica"]).strip() if pd.notna(row["Ulica"]) else ""
    house = str(row["Numer domu"]).strip() if pd.notna(row["Numer domu"]) else ""
    flat = str(row["Numer lokalu"]).strip() if pd.notna(row["Numer lokalu"]) else ""
    number = house if not flat else f"{house}/{flat}"
    kod = str(row["Kod pocztowy"]).strip() if pd.notna(row["Kod pocztowy"]) else ""
    poczta = str(row["Poczta"]).strip() if pd.notna(row["Poczta"]) else ""
    addr_parts = [street, number]
    addr = " ".join(part for part in addr_parts if part)
    if kod or poczta:
        addr = (addr + " - " if addr else "") + " ".join(
            part for part in [kod, poczta] if part
        )
    return addr


def classify_child(type_name: str) -> str:
    """Określa czy składnik to szkoła, przedszkole czy inny typ."""
    text = str(type_name).lower()
    if "przedszko" in text:
        return "przedszkole"
    if "szko" in text:
        return "szkola"
    return "inne"


def main() -> None:
    source_path = next(Path("pobrane").glob("Wykaz_szko*2024_.xlsx"))
    output_path = Path("raporty") / "zespoly_szkolno_przedszkolne_analiza.xlsx"

    df = pd.read_excel(source_path)
    df["adres"] = df.apply(build_address, axis=1)

    parent_mask = (
        df["Rodzaj szkoły/placówki"].eq("jednostka złożona")
        & df["Nazwa placówki"].str.contains("szkolno", case=False, na=False)
        & df["Nazwa placówki"].str.contains("przedszko", case=False, na=False)
    )
    parents = df.loc[parent_mask].copy()
    parent_ids = parents["idPodmiotGlowny"].dropna().unique()

    children = df[df["idPodmiotNadrzedny"].isin(parent_ids)].copy()
    for frame in (parents, children):
        frame["ucz_ogolem"] = pd.to_numeric(frame["ucz_ogolem"], errors="coerce")

    children["kategoria_powiazania"] = children["Typ podmiotu"].apply(classify_child)

    suma_lacznie = children.groupby("idPodmiotNadrzedny")["ucz_ogolem"].sum(
        min_count=1
    )
    suma_szkoly = children.loc[
        children["kategoria_powiazania"] == "szkola"
    ].groupby("idPodmiotNadrzedny")["ucz_ogolem"].sum(min_count=1)
    suma_przedszkola = children.loc[
        children["kategoria_powiazania"] == "przedszkole"
    ].groupby("idPodmiotNadrzedny")["ucz_ogolem"].sum(min_count=1)
    liczba_skladnikow = children.groupby("idPodmiotNadrzedny").size()

    def join_names(frame: pd.DataFrame) -> pd.Series:
        return frame.groupby("idPodmiotNadrzedny")["Nazwa placówki"].apply(
            lambda s: " | ".join(pd.unique(s.dropna()))
        )

    nazwy_szkol = join_names(children.loc[children["kategoria_powiazania"] == "szkola"])
    nazwy_przedszkoli = join_names(
        children.loc[children["kategoria_powiazania"] == "przedszkole"]
    )

    parents["liczba_skladnikow"] = (
        parents["idPodmiotGlowny"].map(liczba_skladnikow).fillna(0).astype(int)
    )
    parents["dzieci_szkola"] = parents["idPodmiotGlowny"].map(suma_szkoly).fillna(0)
    parents["dzieci_przedszkole"] = (
        parents["idPodmiotGlowny"].map(suma_przedszkola).fillna(0)
    )
    parents["powiazana_szkola"] = parents["idPodmiotGlowny"].map(nazwy_szkol).fillna("")
    parents["powiazane_przedszkole"] = (
        parents["idPodmiotGlowny"].map(nazwy_przedszkoli).fillna("")
    )
    parents["dzieci_wyliczone_wartosc"] = (
        parents["dzieci_szkola"] + parents["dzieci_przedszkole"]
    )
    parents["dzieci_z_danych"] = parents["ucz_ogolem"]

    summary_cols = [
        "idPodmiotGlowny",
        "Nazwa placówki",
        "Miejscowość",
        "adres",
        "Typ podmiotu",
        "liczba_skladnikow",
        "powiazana_szkola",
        "dzieci_szkola",
        "powiazane_przedszkole",
        "dzieci_przedszkole",
        "dzieci_wyliczone",  # formuła Excel
        "dzieci_wyliczone_wartosc",
        "dzieci_z_danych",
    ]

    summary = (
        parents.sort_values("dzieci_wyliczone_wartosc", ascending=False)
        .reset_index(drop=True)
        .copy()
    )
    # wypełniamy kolumnę formułą, odwołując się do kolumn z wejściami (H i J)
    summary["dzieci_wyliczone"] = [
        f"=SUM(H{row},J{row})" for row in range(2, len(summary) + 2)
    ]
    summary = summary[summary_cols]

    parents["typ_wiersza"] = "zespol"
    parents["kategoria_powiazania"] = ""
    parents["dzieci_wyliczone"] = parents["dzieci_wyliczone_wartosc"]
    children["typ_wiersza"] = "skladnik"
    children["dzieci_wyliczone"] = children["ucz_ogolem"]
    children["dzieci_z_danych"] = children["ucz_ogolem"]

    details_cols = [
        "typ_wiersza",
        "kategoria_powiazania",
        "idPodmiotGlowny",
        "idPodmiotNadrzedny",
        "Nazwa placówki",
        "Miejscowość",
        "adres",
        "Typ podmiotu",
        "Rodzaj szkoły/placówki",
        "dzieci_z_danych",
        "dzieci_wyliczone",
    ]
    details = pd.concat([parents, children], ignore_index=True)[details_cols]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="podsumowanie_zespolow")
        details.to_excel(writer, index=False, sheet_name="szczegoly_zrodlo")

    print(f"Zapisano raport: {output_path}")
    print("Top 5 zespołów (dzieci_wyliczone_wartosc):")
    print(
        summary[
            ["Nazwa placówki", "Miejscowość", "dzieci_wyliczone_wartosc"]
        ].head(5)
    )


if __name__ == "__main__":
    main()
