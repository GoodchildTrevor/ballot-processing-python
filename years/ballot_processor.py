import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import shutil
import os


def delete_non_breaking_spaces(text):
    if pd.isna(text): return ""
    return str(text).replace("\u00A0", " ").strip()

def postfix(score: int):
    if score == 1:
        return f"{score} балл"
    elif 2 <= score <= 4:
        return f"{score} балла"
    else:
        return f"{score} баллов"

def process_data(df: pd.DataFrame, nominees_df: pd.DataFrame, nominations: list):
    df = df.fillna("xxx")
    
    voter_names = df.iloc[:, 0].astype(str).tolist()
    
    df_votes = df.iloc[:, 1:]
    
    movies_dict = {}
    df_movies = df_votes.iloc[:, 0:10]
    
    for voter_idx, row in enumerate(df_movies.values):
        voter_name = voter_names[voter_idx]
        for pos in range(10):
            movie_key = str(row[pos])
            score = 10 - pos 
            
            if movie_key not in movies_dict:
                movies_dict[movie_key] = {"score": 0, "mentions": []}
            
            movies_dict[movie_key]["score"] += score
            movies_dict[movie_key]["mentions"].append(f"{voter_name} ({postfix(score)})")

    nominations_dict = {nom: {} for nom in nominations}
    df_other = df_votes.iloc[:, 10:]

    for i, nom_cat in enumerate(nominations):
        if i < len(nominees_df.columns):
            current_nominees = nominees_df.iloc[:, i].dropna().unique()
            
            for nominee in current_nominees:
                clean_nominee = delete_non_breaking_spaces(nominee)
                if not clean_nominee: continue
                
                for voter_idx, vote_val in enumerate(df_other.iloc[:, i].values):
                    if clean_nominee in str(vote_val):
                        if clean_nominee not in nominations_dict[nom_cat]:
                            nominations_dict[nom_cat][clean_nominee] = {"score": 0, "mentions": []}
                        
                        nominations_dict[nom_cat][clean_nominee]["score"] += 1
                        nominations_dict[nom_cat][clean_nominee]["mentions"].append(voter_names[voter_idx])

    return movies_dict, nominations_dict

def process_coincidences(nominations_dict):
    main_cats = ["actor", "actress"]
    supp_cats = ["actor2", "actress2"]
    coincidences = []

    for m, s in zip(main_cats, supp_cats):
        data_main = nominations_dict.get(m, {})
        data_supp = nominations_dict.get(s, {})
        
        for name, info in data_main.items():
            if name in data_supp:
                total_score = info["score"] + data_supp[name]["score"]
                combined_mentions = info["mentions"] + data_supp[name]["mentions"]
                coincidences.append((
                    name, total_score, info["score"], data_supp[name]["score"], 
                    ", ".join(combined_mentions)
                ))
    return coincidences

def run_voting(file_path, nominations):
    df_original = pd.read_excel(file_path, sheet_name="номинанты")
    nominees_raw = pd.read_excel(file_path, sheet_name="списки")
    
    nominees_df = nominees_raw.copy()
    nominees_df.columns = nominations[:len(nominees_df.columns)]

    n_votes = len(df_original)
    max_slice = (n_votes // 10) * 10
    slices = list(range(10, max_slice + 1, 10)) if max_slice >= 10 else []

    final_movies, final_noms = process_data(df_original, nominees_df, nominations)
    
    results_map = [("победители", (final_movies, final_noms))]
    for s in slices:
        subset = df_original.iloc[:s, :]
        results_map.append((f"победители {s}", process_data(subset, nominees_df, nominations)))
    
    wb = openpyxl.load_workbook(file_path)
    
    for sheet_name, (m_dict, n_dict) in results_map:
        if sheet_name in wb.sheetnames: del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        
        sorted_movies = sorted([k for k in m_dict if k != "xxx"], 
                               key=lambda x: m_dict[x]["score"], reverse=True)
        
        headers = ["movie"] + nominations
        for z, h in enumerate(headers):
            col = z * 3 + 1
            ws.cell(1, col, h)
            ws.cell(1, col + 1, f"points_{h}")
            ws.cell(1, col + 2, f"mentions_by_{h}")

        for row_idx, movie in enumerate(sorted_movies, start=2):
            ws.cell(row_idx, 1, movie)
            ws.cell(row_idx, 2, m_dict[movie]["score"])
            ws.cell(row_idx, 3, ", ".join(m_dict[movie]["mentions"]))

        for col_idx, nom in enumerate(nominations):
            start_col = 4 + (col_idx * 3)
            sorted_noms = sorted(n_dict[nom].items(), key=lambda x: x[1]["score"], reverse=True)
            
            for row_offset, (name, info) in enumerate(sorted_noms, start=2):
                ws.cell(row_offset, start_col, name)
                ws.cell(row_offset, start_col + 1, info["score"])
                ws.cell(row_offset, start_col + 2, ", ".join(info["mentions"]))

    sn_coinc = "совпадения"
    if sn_coinc in wb.sheetnames: del wb[sn_coinc]
    ws_c = wb.create_sheet(sn_coinc)
    ws_c.append(["Имя", "Баллы (итого)", "Первый план", "Второй план", "Упоминания"])
    
    coincidences = process_coincidences(final_noms)
    for row in coincidences:
        ws_c.append(row)

    wb.save(file_path)