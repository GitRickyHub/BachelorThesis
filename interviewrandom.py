import os
import random
from pathlib import Path


folder_path = # Insert desired path

# Collect all .pdf files (sorted for consistency)
pdf_files = sorted([f.name for f in folder_path.glob("*.pdf")])
n_files = len(pdf_files)

if n_files == 0:
    print("No .pdf files found in the folder.")
    exit()

# Define anonymization tools
tools = ["Mistral", "Llama", "Phi", "Textwash"]

# Calculate fair distribution
base_count = n_files // len(tools)
remainder = n_files % len(tools)

assignment_counts = [base_count] * len(tools)
for i in range(remainder):
    assignment_counts[i] += 1

# Randomize which tool gets extra files
tool_distribution = list(zip(tools, assignment_counts))
random.shuffle(tool_distribution)

# Create tool list with correct counts
assignment_list = []
for tool, count in tool_distribution:
    assignment_list.extend([tool] * count)
random.shuffle(assignment_list)

# Assign each PDF file a tool
file_tool_mapping = list(zip(pdf_files, assignment_list))

# Print the results
for filename, tool in file_tool_mapping:
    print(f"{filename} -> {tool}")
