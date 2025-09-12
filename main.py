import pandas as pd
import openpyxl

def load_csv(file_path_input: str, file_path_pay:str) -> pd.DataFrame:
    """CSVを読み込む"""
    return pd.read_csv(file_path_input), pd.read_csv(file_path_pay)

def clean_data(df: pd.DataFrame, payouts_df) -> pd.DataFrame:
    """データの整形（着順を整数に変換）"""
    df['着順'] = df['着順'].astype(int)

    # 払戻金の整形（数字だけに変換）
    payouts_df["払戻金"] = (
        payouts_df["払戻金"]
        .str.replace("円", "", regex=False)  # 円を削除
        .str.replace(",", "", regex=False)   # カンマを削除
        .astype(int)                         # 数値化
    )

    return df, payouts_df

def calc_summary(df_tm, df_payouts, trackmen=['TM1', 'TM2', 'TM3', 'TM4'],
                  bet_types=["単勝","複勝"]):
    
    results = []
    # 開催場ごとの集計
    for venue in df_tm["開催場"].unique():
        venue_races = df_tm[df_tm["開催場"] == venue]
        venue_race_ids = venue_races["レースID"].unique()

    #レースごとに的中判定
        for race_id in venue_race_ids:
            race_horses = venue_races[venue_races["レースID"] == race_id]
            payouts = df_payouts[df_payouts["レースID"] == race_id]

            sorted_horses = race_horses.sort_values("着順")
            winner_row = sorted_horses.iloc[0]
            winner = str(winner_row["馬番"])
            top2 = sorted_horses.head(2)["馬番"].astype(str).tolist()
            top3 = sorted_horses.head(3)["馬番"].astype(str).tolist() # 着順上位3頭->3着以内

            # トラックマンごとの判定
            for t in trackmen:
                horses = race_horses.loc[race_horses[t]=="◎", "馬番"].astype(str).tolist()
                if not horses:
                    continue
                other_marks = ["○","▲","△","穴"]
                other_horses = race_horses.loc[race_horses[t].isin(other_marks), "馬番"].astype(str).tolist()

                # 券種ごとの判定
                for bet in bet_types:
                    hit = 0
                    hits_count = 0  # 的中回数
                    pay = 0 # pay:購入金額

                    if bet == "単勝":
                        pay = len(horses) * 100
                        for h in horses:
                            if h == winner:
                                hit += payouts.loc[(payouts["券種"]=="単勝") & (payouts["組番"]==h), "払戻金"].sum()
                                hits_count += 1

                    elif bet == "複勝":
                        pay = len(horses) * 100
                        for h in horses:
                            if h in top3:
                                hit += payouts.loc[(payouts["券種"]=="複勝") & (payouts["組番"]==h), "払戻金"].sum()
                                hits_count += 1

                
                    results.append({
                        "開催場": venue,
                        "レースID": race_id,
                        "トラックマン": t,
                        "券種": bet,
                        "購入金額": pay,
                        "払戻金": hit,
                        "収支": hit - pay,
                        "回収率(%)": round(hit/pay*100,1) if pay>0 else None,
                    })

    return pd.DataFrame(results)

def export_to_excel(df: pd.DataFrame, output_path: str):
    """Excelに出力"""
    df.to_excel(output_path, index=False)
    print(f"Excelに出力したよ！ -> {output_path}")

def aggregate_data(df: pd.DataFrame) -> pd.DataFrame:
    """トラックマンごと＆券種ごとの集計"""
    grouped = df.groupby(["トラックマン", "券種"]).agg({
        "購入金額": "sum",
        "払戻金": "sum",
        "収支": "sum"
    }).reset_index()

    # 回収率を新しく計算（払戻金 ÷ 購入金額 × 100）
    grouped["回収率(%)"] = (grouped["払戻金"] / grouped["購入金額"] * 100).round(1)

    return grouped


def main():
    input_csv = "input.csv"   # 元データ
    payouts_csv = 'payouts.csv' # 払戻データ
    output_excel = "output.xlsx"  # 出力先

    df, payouts_df = load_csv(input_csv, payouts_csv)
    df, payouts_df = clean_data(df, payouts_df)
    df = calc_summary(df, payouts_df)
    df = aggregate_data(df)
    export_to_excel(df, output_excel)

if __name__ == "__main__":
    main()

