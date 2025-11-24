import pandas as pd
from pathlib import Path

BASE_DIR = Path("pobrane/GUS")
OUTPUT_XLSX = Path("raporty/demografia_dzieci.xlsx")

# Ścieżki do plików źródłowych
POWIAT_FILE = BASE_DIR / "2023_2060_4-powiaty" / "Powiaty" / "24 ÿlÑskie" / "2411 raciborski.xlsx"
MIASTO_FILE = (
    BASE_DIR
    / "2023_2040_8_prognoza_ludnosci_dla_gmin_na_lata_2023-2040_2"
    / "24 śląskie"
    / "2411 powiat raciborski"
    / "2411011 Racibórz (M).xlsx"
)
TABLICA_ZBIORCZA = BASE_DIR / "2023_2040_9_gminy_ludnosc_-_tablica_zbiorcza_2.xlsx"


def group_age(age: int) -> str:
    if 0 <= age <= 2:
        return "zlobek_0_2"
    if 3 <= age <= 6:
        return "przedszkole_3_6"
    if 7 <= age <= 18:
        return "szkolne_7_18"
    return "poza_zakresem"


def load_powiat():
    df = pd.read_excel(POWIAT_FILE, sheet_name="Tabl. 1", skiprows=6)
    df = df[df["Wiek Age"].apply(lambda x: pd.api.types.is_number(x))]
    # szeroki -> długi
    year_cols = [c for c in df.columns if isinstance(c, (int, float)) or str(c).startswith("202")]
    melted = df.melt(id_vars=["Wiek Age"], value_vars=year_cols, var_name="rok", value_name="liczba")
    # normalizujemy nagłówek 2022*
    melted["rok"] = melted["rok"].replace({"2022*": 2022})
    melted = melted[melted["liczba"].notna()]
    melted["wiek"] = melted["Wiek Age"].astype(int)
    melted["rok"] = melted["rok"].astype(int)
    melted["grupa"] = melted["wiek"].apply(group_age)
    agg = (
        melted[melted["grupa"] != "poza_zakresem"]
        .groupby(["rok", "grupa"], as_index=False)["liczba"]
        .sum()
    )
    agg["jednostka"] = "Powiat raciborski"
    agg["typ"] = "powiat"
    agg["uwaga"] = "Dokładne wartości z Tablica 1 (jednoroczne wieki 0-100)"
    return agg[["jednostka", "typ", "rok", "grupa", "liczba", "uwaga"]]


def load_miasto():
    # Dane z pliku gminnego (zakres 0-9, 10-19)
    df = pd.read_excel(MIASTO_FILE, sheet_name="Tabl. 1", skiprows=6)
    child_rows = df[df["Grupa wieku   Age group"].isin(["0-9", "10-19", "Ogółem Total"])]
    year_cols = [c for c in child_rows.columns if isinstance(c, (int, float)) or str(c).startswith("202")]
    melted = child_rows.melt(
        id_vars=["Grupa wieku   Age group"], value_vars=year_cols, var_name="rok", value_name="liczba"
    )
    melted["rok"] = melted["rok"].replace({"2022*": 2022})
    melted = melted[melted["liczba"].notna()]
    melted["rok"] = melted["rok"].astype(int)
    # Mapowanie na nasze grupy (uwaga: brak rozbicia na 0-2 i 3-6)
    def map_group(label: str) -> str:
        if label == "0-9":
            return "dzieci_0_9 (brak rozbicia 0-2/3-6)"
        if label == "10-19":
            return "mlodziez_10_19 (przybliżenie grupy szkolnej)"
        if label == "Ogółem Total":
            return "ogolem"
        return label

    melted["grupa"] = melted["Grupa wieku   Age group"].apply(map_group)
    melted["jednostka"] = "Miasto Racibórz"
    melted["typ"] = "gmina"
    melted["uwaga"] = "Dane dostępne tylko w grupach 0-9 i 10-19 (Tabl.1 prognoza gmin 2023-2040)"

    # Dodaj 0-17 z Tabl. 2 (bliżej definicji wieku szkolnego)
    df2 = pd.read_excel(MIASTO_FILE, sheet_name="Tabl. 2", skiprows=6)
    row_0_17 = df2[df2["Grupa wieku   Age group"] == "0-17"]
    if not row_0_17.empty:
        melted_0_17 = row_0_17.melt(
            id_vars=["Grupa wieku   Age group"],
            value_vars=[c for c in row_0_17.columns if isinstance(c, (int, float)) or str(c).startswith("202")],
            var_name="rok",
            value_name="liczba",
        )
        melted_0_17 = melted_0_17[melted_0_17["liczba"].notna()]
        melted_0_17["rok"] = melted_0_17["rok"].replace({"2022*": 2022})
        melted_0_17["rok"] = melted_0_17["rok"].astype(int)
        melted_0_17["grupa"] = "dzieci_0_17 (brak rozbicia na 0-2/3-6/7-17)"
        melted_0_17["jednostka"] = "Miasto Racibórz"
        melted_0_17["typ"] = "gmina"
        melted_0_17["uwaga"] = "Zakres 0-17 z Tabl.2 (prognoza gmin 2023-2040)"
        melted = pd.concat([melted, melted_0_17], ignore_index=True)

    # (opcjonalnie) uzupełnij danymi z tablicy zbiorczej jeśli potrzebne
    try:
        zb = pd.read_excel(TABLICA_ZBIORCZA, sheet_name="Tabela zbiorcza")
        zb_rac = zb[(zb["Kod_TERYT"] == 2411011) & (zb["Płeć"] == "Ogółem") & (zb["Wiek"] == "0-17")]
        if not zb_rac.empty:
            zb_melt = zb_rac.melt(
                id_vars=["Wiek"], value_vars=[c for c in zb_rac.columns if isinstance(c, (int, float)) or str(c).startswith("202")],
                var_name="rok", value_name="liczba"
            )
            zb_melt["rok"] = zb_melt["rok"].astype(int)
            zb_melt["grupa"] = "dzieci_0_17 (tablica zbiorcza)"
            zb_melt["jednostka"] = "Miasto Racibórz"
            zb_melt["typ"] = "gmina"
            zb_melt["uwaga"] = "Dane z tablicy zbiorczej gmin (0-17)"
            melted = pd.concat([melted, zb_melt], ignore_index=True)
    except FileNotFoundError:
        pass

    return melted[["jednostka", "typ", "rok", "grupa", "liczba", "uwaga"]]


def main():
    powiat = load_powiat()
    miasto = load_miasto()

    # Zbiorcza tabela (long)
    combined = pd.concat([powiat, miasto], ignore_index=True)
    combined = combined[
        ["jednostka", "typ", "rok", "grupa", "liczba", "uwaga"]
    ].sort_values(["jednostka", "rok", "grupa"])

    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        powiat.to_excel(writer, sheet_name="powiat_raciborski", index=False)
        miasto.to_excel(writer, sheet_name="miasto_raciborz", index=False)
        combined.to_excel(writer, sheet_name="zestawienie", index=False)

    print(f"Zapisano {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
