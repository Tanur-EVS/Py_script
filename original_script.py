import os
import re
import pandas as pd
 
def clean_field(field):
    return field.strip().strip('[]')
 
function_map = [
 
    # String functions
    (r'\bREPLACE\((.*?),\s*(.*?),\s*(.*?)\)', lambda m: f'SUBSTITUTE({m.group(1)}, {m.group(2)}, {m.group(3)})'),
    (r'\bFIND\((.*?),\s*(.*?)(?:,\s*(.*?))?\)', lambda m: f'SEARCH({m.group(1)}, {m.group(2)})'),
    (r'\bCONTAINS\((.*?)\,\s*(.*?)\)', r'SEARCH(\2, \1, 1, 0) > 0'),
    (r'\bSTARTSWITH\((.*?),\s*(.*?)\)', r'LEFT(\1, LEN(\2)) = \2'),
    (r'\bENDSWITH\((.*?),\s*(.*?)\)', r'RIGHT(\1, LEN(\2)) = \2'),
    (r'\bSPLIT\((.*?)\,\s*"(.*?)"(?:\,\s*(\d+))?\)', lambda m: f'PATHITEM(SUBSTITUTE({m.group(1)}, "{m.group(2)}", "|"), {m.group(3) or 1})'),
    # Date functions
    (r'\bDATEPART\(\s*[\'"]year[\'"]\s*,\s*(.*?)\)', r'YEAR(\1)'),
    (r'\bDATEPART\(\s*[\'"]month[\'"]\s*,\s*(.*?)\)', r'MONTH(\1)'),
    (r'\bDATEPART\(\s*[\'"]day[\'"]\s*,\s*(.*?)\)', r'DAY(\1)'),
    (r'\bDATEPART\(\s*[\'"]weekday[\'"]\s*,\s*(.*?)\)', r'WEEKDAY(\1)'),
    (r'\bTODAY\(\)', r'TODAY()'),
    (r'\bNOW\(\)', r'NOW()'),
    (r'\bMAKEDATE\((\d+),\s*(\d+),\s*(\d+)\)', lambda m: f'DATE({m.group(1)}, {m.group(2)}, {m.group(3)})'),
    (r'\bMAKETIME\((.*?),\s*(.*?),\s*(.*?)\)', r'TIME(\1, \2, \3)'),
    (r'\bDATEADD\((.*?),\s*(.*?),\s*[\'"](\w+)[\'"]\)', lambda m: f'DATEADD({m.group(1)}, {m.group(2)}, {m.group(3).upper()})'),
    (r'\bDATEDIFF\((.*?),\s*(.*?),\s*[\'"](\w+)[\'"]\)', lambda m: f'DATEDIFF({m.group(1)}, {m.group(2)}, {m.group(3).upper()})'),
    (r'\bDATETRUNC\(\s*[\'"](year|month|day)[\'"]\s*,\s*(.*?)\)',
        lambda m: (
            f'DATE(YEAR({m.group(2)}), 1, 1)' if m.group(1).lower() == 'year'
            else f'EOMONTH({m.group(2)}, -1) + 1' if m.group(1).lower() == 'month'
            else f'TRUNC({m.group(2)}, "DAY")'
        )),
    (r'\bDATETRUNC\(\s*[\'"]quarter[\'"]\s*,\s*(.*?)\)',
        lambda m: f'DATE(YEAR({m.group(1)}), ((QUARTER({m.group(1)}) - 1) * 3) + 1, 1)'),
 
    # Logical functions
    (r'\bIF\s+(.*?)\s+THEN\s+(.*?)\s+END', r'IF(\1, \2)'),
    (r'\bIF\s+(.*?)\s+THEN\s+(.*?)\s+ELSE\s+(.*?)\s+END', r'IF(\1, \2, \3)'),
    (r'\bIIF\((.*?),\s*(.*?),\s*(.*?)\)', r'IF(\1, \2, \3)'),
    (r'\bAND\((.*?)\)', lambda m: " && ".join(map(str.strip, m.group(1).split(',')))),
    (r'\bOR\((.*?)\)', lambda m: " || ".join(map(str.strip, m.group(1).split(',')))),
    (r'\bNOT\((.*?)\)', lambda m: f'NOT({m.group(1)})'),

    (r'\bWINDOW_SUM\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(SUM({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_AVG\(AVG\((.*?)\)\)', lambda m: f'CALCULATE(AVERAGE({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_AVG\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(AVERAGEX(ALL(Table), {m.group(1)}))'),
    (r'\bWINDOW_SUM\(AVG\((.*?)\)\)', lambda m: f'CALCULATE(SUMX(ALL(Table), {m.group(1)}))'),
    (r'\bWINDOW_MAX\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(MAX({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_MIN\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(MIN({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_VAR\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(VAR.S({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_VAR\(AVG\((.*?)\)\)', lambda m: f'CALCULATE(VAR.S({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_STDEV\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(STDEV.S({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_STDEV\(MAX\((.*?)\)\)', lambda m: f'CALCULATE(STDEV.S({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_MIN\(MIN\((.*?)\)\)', lambda m: f'CALCULATE(MIN({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_MAX\(MAX\((.*?)\)\)', lambda m: f'CALCULATE(MAX({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_SUM\(COUNT\((.*?)\)\)', lambda m: f'CALCULATE(COUNT({m.group(1)}), REMOVEFILTERS())'),
    (r'\bWINDOW_AVG\(COUNTD\((.*?)\)\)', lambda m: f'CALCULATE(DISTINCTCOUNT({m.group(1)}), REMOVEFILTERS())'),
 
    # Aggregates and math
    (r'\bSUM\((.*?)\)', r'SUM(\1)'),
    (r'\bAVG\((.*?)\)', r'AVERAGE(\1)'),
    (r'\bMIN\((.*?)\)', r'MIN(\1)'),
    (r'\bMAX\((.*?)\)', r'MAX(\1)'),
    (r'\bSQRT\((.*?)\)', r'SQRT(\1)'),
    (r'\bLOG\((.*?)\)', r'LOG(\1)'),
    (r'\bINT\((.*?)\)', r'INT(\1)'),
    (r'\bEXP\((.*?)\)', r'EXP(\1)'),
    (r'\bZN\((.*?)\)', r'IF(ISBLANK(\1), 0, \1)'),
    (r'\bISNULL\((.*?)\)', r'ISBLANK(\1)'),
    (r'\bIFNULL\((.*?),\s*(.*?)\)', r'IF(ISBLANK(\1), \2, \1)'),
    (r'\bNOT\s+ISNULL\((.*?)\)', r'NOT(ISBLANK(\1))'),
    (r'\bCOUNTD\((.*?)\)', r'DISTINCTCOUNT(\1)'),
 
    # INDEX() → RANKX(ALL(), [Measure]) — you may want to replace [Measure] dynamically
    (r'\bINDEX\(\)', lambda m: f'RANKX(ALL(), [Measure])'),
 
    # PREVIOUS_VALUE([Sales]) → CALCULATE([Sales], PREVIOUSMONTH([Date]))
    (r'\bPREVIOUS_VALUE\((.*?)\)', lambda m: f'CALCULATE({m.group(1)}, PREVIOUSMONTH([Date]))'),
 
    # Running aggregates
    (r'\bRUNNING_SUM\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(SUM({m.group(1)}), FILTER(ALL(Table), Table[Date] <= MAX(Table[Date])))'),
    (r'\bRUNNING_AVG\(SUM\((.*?)\)\)', lambda m: f'CALCULATE(AVERAGEX(FILTER(ALL(Table), Table[Date] <= MAX(Table[Date])), {m.group(1)}))'),
 
    # Lookup and previous value
    (r'\bLOOKUP\(SUM\((.*?)\),\s*(-?\d+)\)', lambda m: f'CALCULATE(SUM({m.group(1)}), DATEADD(Table[Date], {m.group(2)}, MONTH))'),
    (r'\bPREVIOUS_VALUE\((.*?)\)', lambda m: f'VAR Prev = CALCULATE([Sales], DATEADD(Table[Date], -1, MONTH)) RETURN Prev + [Sales] + [Sales]'),
 
    # Dense Rank
    (r'\bRANK_DENSE\(SUM\((.*?)\)\)', lambda m: f"RANKX(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, DENSE)"),
 
    # Simple Rank
    (r'\bRANK\(SUM\((.*?)\)\)', lambda m: f"RANKX(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, SKIP)"),
 
    # Generic Rank
    (r'\bRANK\((.*?)\)', lambda m: f"RANKX(ALL('Table'), {clean_field(m.group(1))})"),
 
    # Unique Rank
    (r'\bRANK_UNIQUE\(SUM\((.*?)\)\)', lambda m: f"RANKX(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, SKIP)  // Add tie-breaker logic if needed"),
 
    # Modified Rank
    (r'\bRANK_MODIFIED\(SUM\((.*?)\)\)', lambda m: f"RANKX(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, SKIP)  // Simulate modified rank manually"),
 
    # Rank with Partition (e.g., by Category)
    (r'\bRANK\(SUM\((.*?)\)\)\s+BY\s+([a-zA-Z_][a-zA-Z0-9_]*)', lambda m: f"RANKX(ALL('Table'[{clean_field(m.group(2))}]), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, SKIP)"),
 
    # Top N Filter
    (r'\bIF\s+RANK\(SUM\((.*?)\)\,\s*\'desc\'\)\s*<=\s*(\d+)', lambda m: f"IF(RANKX(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), BLANK(), DESC, SKIP) <= {m.group(2)}, TRUE(), FALSE())"),
 
    # Percentile Rank (approximation)
    (r'\bWINDOW_PERCENTILE\(SUM\((.*?)\),\s*(\d+)\)', lambda m: f"PERCENTILEX.INC(ALL('Table'), CALCULATE(SUM('Table'[{clean_field(m.group(1))}])), {int(m.group(2))/100})"),
 
    # Percentile Rank (approximation)
    (r'\bWINDOW_PERCENTILE\(SUM\((.*?)\),\s*(\d+)\)', lambda m: f"PERCENTILEX.INC(ALL('Table'), CALCULATE(SUM('Table'[{m.group(1)}])), {int(m.group(2))/100})"),
 
    # Percentile
    (r'\bPERCENTILE\((.*?),\s*(.*?)\)',  lambda m: f"PERCENTILEX.INC(ALL('Table'), 'YourTable'[{m.group(1).strip('[]')}], {m.group(2)})"),
    (r'\bPERCENTILE\((.*?),\s*(.*?)\)', lambda m: f"PERCENTILEX.INC(REMOVEFILTERS(), {clean_field(m.group(1))}, {m.group(2)})")

]

#HelperFunction
def apply_func_mapping(expr: str) -> str:
    for pattern, repl in function_map:
        expr = re.sub(pattern, repl if not callable(repl) else repl, expr, flags=re.IGNORECASE)
    return expr
 
 
def convert_lod_expression(expr, table_name="Table"):
    expr = expr.strip()

    # FIXED LOD
    if expr.upper().startswith("{FIXED"):
        match = re.match(r'\{\s*FIXED\s+([^\:]+?)\s*:\s*(.*?)\}', expr, re.IGNORECASE)
        if match:
            dims = [d.strip().strip("[]") for d in match.group(1).split(",")]
            measure = match.group(2).strip()
            measure=apply_func_mapping(measure)
            dim_refs = ", ".join([f"'{table_name}'[{d}]" for d in dims])
            return f"CALCULATE({measure}, ALLEXCEPT('{table_name}', {dim_refs}))"

    # INCLUDE LOD
    if expr.upper().startswith("{INCLUDE"):
        match = re.match(r'\{\s*INCLUDE\s+([^\:]+?)\s*:\s*(.*?)\}', expr, re.IGNORECASE)
        if match:
            dims = [d.strip().strip(" []") for d in match.group(1).split(",")]
            measure = match.group(2).strip()
            measure = apply_func_mapping(measure)
            filters = ", ".join([f"KEEPFILTERS(VALUES('{table_name}'[{d}]))" for d in dims])
            return f"CALCULATE({measure}, {filters})"

    # EXCLUDE LOD
    if expr.upper().startswith("{EXCLUDE"):
        match = re.match(r'\{\s*EXCLUDE\s+([^\:]+?)\s*:\s*(.*?)\}', expr, re.IGNORECASE)
        if match:
            dims = [d.strip().strip("[]") for d in match.group(1).split(",")]
            measure = match.group(2).strip()
            measure = apply_func_mapping(measure)
            filters = ", ".join([f"REMOVEFILTERS('{table_name}'[{d}])" for d in dims])
            return f"CALCULATE({measure}, {filters})"

    # If no match, return original expression
    return expr

 
def multi_if_to_switch(expr):
    expr = expr.strip()
    if not expr.upper().startswith("IF"):
        return expr
    expr = re.sub(r'\s+', ' ', expr)
    else_part = None
    if re.search(r'\bELSE\b', expr, re.IGNORECASE):
        split_els = re.split(r'\bELSE\b', expr, flags=re.IGNORECASE)
        if len(split_els) > 1:
            else_segment = split_els[1]
            else_part = else_segment.replace("END", "").strip()
        expr = split_els[0]
    conditions = []
    results = []
    parts = re.split(r'\bELSEIF\b', expr, flags=re.IGNORECASE)
    first_if = parts[0].strip()
    if_match = re.match(r'IF\s+(.*?)\s+THEN\s+(.*)', first_if, flags=re.IGNORECASE)
    if if_match:
        conditions.append(if_match.group(1).strip())
        results.append(if_match.group(2).strip())
    for p in parts[1:]:
        if_match = re.match(r'(.*?)\s+THEN\s+(.*)', p.strip(), flags=re.IGNORECASE)
        if if_match:
            conditions.append(if_match.group(1).strip())
            results.append(if_match.group(2).strip())
    switch_parts = [f"{cond}, {res}" for cond, res in zip(conditions, results)]
    switch_str = f"SWITCH(TRUE(), {', '.join(switch_parts)}"
    if else_part:
        switch_str += f", {else_part}"
    switch_str += ")"
    return switch_str
 
def convert_tableau_to_dax(expr, table_name="Table"):
    if pd.isna(expr):
        return None
    expr = str(expr).strip()
    # Handle LOD expressions
    if expr.startswith("{"):
        return convert_lod_expression(expr, table_name=table_name)
    
    # Multi-condition IF
 
    if re.search(r'\bELSEIF\b', expr, re.IGNORECASE):
        return multi_if_to_switch(expr)
 
    # CASE to SWITCH (simple)
    if expr.upper().startswith("CASE"):
        case_matches = re.findall(r'WHEN (.*?) THEN (.*?)(?= WHEN| ELSE| END)', expr, re.IGNORECASE)
        else_match = re.search(r'ELSE (.*?) END', expr, re.IGNORECASE)
        switch_parts = []
        for cond, result in case_matches:
            switch_parts.append(f"{cond}, {result}")
        if else_match:
            switch_str = f"SWITCH(TRUE(), {', '.join(switch_parts)}, {else_match.group(1)})"
        else:
            switch_str = f"SWITCH(TRUE(), {', '.join(switch_parts)})"
        return switch_str
 
    # Regex replacements
    result = expr
    for pattern, replacement in function_map:
        if callable(replacement):
            result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
        else:
            result = re.sub(pattern, replacement, result, flags=re.IGNORECASE)
    return result
 
def process_all_excels_pure(input_folder, output_folder, colname="calculation"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xlsx')]
    for filename in excel_files:
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, filename)
        xl = pd.ExcelFile(input_path)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                col_candidates = [c for c in df.columns if c.strip().lower() == colname.strip().lower()]
                if col_candidates:
                    calc_col = col_candidates[0]
                    temp_df = pd.DataFrame()
                    temp_df["calculation"] = df[calc_col]
                    temp_df["DAX Expressions"] = df[calc_col].apply(convert_tableau_to_dax)
                else:
                    temp_df = pd.DataFrame(columns=["calculation", "DAX Expressions"])
                temp_df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"{filename} converted and saved to {output_path}")
 
if __name__ == "__main__":
    input_folder = r"C:\Users\Tanur.Yadav\Desktop\tab to bi\Sample_Queries"
    output_folder = r"C:\Users\Tanur.Yadav\Desktop\tab to bi\Sample_Queries_Output"
    process_all_excels_pure(input_folder, output_folder)
 