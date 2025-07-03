
import os
import pandas as pd
from collections import defaultdict

# === CONFIG ===
input_folder = # Insert desired input path
summary_output = os.path.join(input_folder, "NER_Entity_Summary.xlsx")

# === Collect all entity scores ===
entity_scores = defaultdict(list)
micro_scores = {'Precision': [], 'Recall': [], 'F1': []}

for file in os.listdir(input_folder):
    if not file.endswith('.xlsx'):
        continue

    file_path = os.path.join(input_folder, file)

    try:
        df = pd.read_excel(file_path, sheet_name='Entity_Metrics')
    except Exception as e:
        print(f"[SKIPPED] {file} — {e}")
        continue

    for _, row in df.iterrows():
        entity = row['Entity']
        if entity == 'MICRO_AVG':
            for metric in ['Precision', 'Recall', 'F1']:
                micro_scores[metric].append(row[metric])
        elif entity not in ['MACRO_AVG']:
            for metric in ['Precision', 'Recall', 'F1']:
                entity_scores[entity].append((metric, row[metric]))

# === Compute averages per entity ===
aggregated = []
for entity, values in entity_scores.items():
    precision_vals = [v for m, v in values if m == 'Precision']
    recall_vals = [v for m, v in values if m == 'Recall']
    f1_vals = [v for m, v in values if m == 'F1']

    aggregated.append({
        'Entity': entity,
        'Average Precision': sum(precision_vals) / len(precision_vals) if precision_vals else 0,
        'Average Recall': sum(recall_vals) / len(recall_vals) if recall_vals else 0,
        'Average F1': sum(f1_vals) / len(f1_vals) if f1_vals else 0,
        'Files Count': len(precision_vals)
    })

# === Compute global micro averages ===
aggregated.append({
    'Entity': 'GLOBAL_MICRO_AVG',
    'Average Precision': sum(micro_scores['Precision']) / len(micro_scores['Precision']) if micro_scores['Precision'] else 0,
    'Average Recall': sum(micro_scores['Recall']) / len(micro_scores['Recall']) if micro_scores['Recall'] else 0,
    'Average F1': sum(micro_scores['F1']) / len(micro_scores['F1']) if micro_scores['F1'] else 0,
    'Files Count': len(micro_scores['F1'])
})

# === Save summary to Excel ===
summary_df = pd.DataFrame(aggregated)
summary_df = summary_df.sort_values(by='Entity')

with pd.ExcelWriter(summary_output, engine='xlsxwriter') as writer:
    summary_df.to_excel(writer, index=False, sheet_name='Summary')

print(f"\n✅ Summary saved to:\n{summary_output}")
