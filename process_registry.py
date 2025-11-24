import pandas as pd
from pathlib import Path

REGISTRY_FILE = Path("pobrane/Wykaz_szkół_i_placówek_oświatowych_30.09.2024_.xlsx")
OUT_FILE = Path("raporty/placowki_registry.xlsx")

POWIAT_FILTER = "raciborsk"
MIASTO_FILTER = "Racibórz"


def classify_kind(typ_podmiotu: str) -> str:
    """Sprowadź typ podmiotu do kilku kategorii użytecznych dla analiz."""
    if not typ_podmiotu:
        return "inne"
    name = typ_podmiotu.lower()
    if "żłob" in name or "zlob" in name:
        return "zlobek"
    if "przedszk" in name or "punkt przedszkolny" in name:
        return "przedszkole"
    if "szkoła podstawowa" in name or "szkola podstawowa" in name:
        return "szkola_podstawowa"
    if "liceum" in name or "technikum" in name or "branż" in name or "policealn" in name:
        return "szkola_ponadpodstawowa"
    if "zespół szkół" in name or "zespół szk" in name or "zespól" in name:
        return "zespol_szkol"
    return "inne"


def load_registry():
    df = pd.read_excel(REGISTRY_FILE)
    df["Rodzaj_kategorii"] = df["Typ podmiotu"].apply(classify_kind)
    return df


def summarize(df: pd.DataFrame, label: str) -> dict:
    fields_sum = ["ucz_ogolem", "w tym_ucz_dziewczeta", "w tym_w oddz_przedszk", "lb_oddz"]
    agg = (
        df.groupby("Rodzaj_kategorii")
        .agg(
            liczba_placowek=("Nazwa placówki", "count"),
            **{f"sum_{f}": (f, "sum") for f in fields_sum if f in df.columns},
        )
        .reset_index()
    )
    agg.insert(0, "obszar", label)
    return agg


def main():
    df = load_registry()
    # powiat
    df_pow = df[df["Powiat"].str.contains(POWIAT_FILTER, case=False, na=False)].copy()
    # miasto
    df_miasto = df_pow[df_pow["Gmina"].str.contains(MIASTO_FILTER, case=False, na=False)].copy()

    sum_pow = summarize(df_pow, "Powiat raciborski")
    sum_miasto = summarize(df_miasto, "Miasto Racibórz")

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        df_pow.to_excel(writer, sheet_name="powiat_raciborski_detailed", index=False)
        df_miasto.to_excel(writer, sheet_name="miasto_raciborz_detailed", index=False)
        sum_pow.to_excel(writer, sheet_name="powiat_raciborski_podsumowanie", index=False)
        sum_miasto.to_excel(writer, sheet_name="miasto_raciborz_podsumowanie", index=False)

    print(f"Zapisano {OUT_FILE}")


if __name__ == "__main__":
    main()
