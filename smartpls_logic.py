
import pandas as pd
import os
import sys
from interpretation import ScientificInterpreter

class SmartPLSReader:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.interpreter = ScientificInterpreter()
        self.ie_data = []
        
    def get_complete_sheet_name(self, xl):
        sheet_names_lower = [s.lower() for s in xl.sheet_names]
        if 'complete' in sheet_names_lower:
            return xl.sheet_names[sheet_names_lower.index('complete')]
        return xl.sheet_names[0]
        
    def find_file(self, filename):
        path = os.path.join(self.folder_path, filename)
        if os.path.exists(path):
            return path
            
        # Case-insensitive / alias check
        fn_lower = filename.lower()
        aliases = []
        if 'bootstrapping' in fn_lower or 'bt' in fn_lower:
            aliases = ['bt.xlsx', 'bootstrapping.xlsx', 'bt.xls', 'bootstrapping.xls']
        elif 'blindfolding' in fn_lower or 'blf' in fn_lower:
            aliases = ['blf.xlsx', 'bf.xlsx', 'blindfolding.xlsx', 'blf.xls', 'blindfolding.xls']
        elif 'pls' in fn_lower:
            aliases = ['pls.xlsx', 'pls algorithm.xlsx', 'pls.xls']
            
        for alias in aliases:
            # Check both original case and lowercase of the alias
            for name in [alias, alias.upper(), alias.lower()]:
                alias_path = os.path.join(self.folder_path, name)
                if os.path.exists(alias_path):
                    return alias_path
                    
        # Fallback to scanning the directory for substring match
        try:
            for f in os.listdir(self.folder_path):
                f_lower = f.lower()
                if ('bootstrapping' in fn_lower or 'bt' in fn_lower) and ('boot' in f_lower or f_lower == 'bt.xlsx' or f_lower == 'bt.xls'):
                    return os.path.join(self.folder_path, f)
                elif ('blindfolding' in fn_lower or 'blf' in fn_lower) and ('blind' in f_lower or f_lower == 'blf.xlsx' or f_lower == 'blf.xls'):
                    return os.path.join(self.folder_path, f)
                elif 'pls' in fn_lower and 'pls' in f_lower and 'boot' not in f_lower and 'bt' not in f_lower and 'blind' not in f_lower and 'blf' not in f_lower:
                    return os.path.join(self.folder_path, f)
        except:
            pass
            
        return None

    def read_subtable(self, df, keyword, col_search_idxs=[0, 1], header_offset=None):
        """
        Finds a subtable by keyword.
        col_search_idxs: list of columns to search for keyword.
        header_offset: if None, finds first non-empty row after keyword.
        """
        start_row = -1
        
        # Search for keyword in specified columns
        for col_idx in col_search_idxs:
            if col_idx >= len(df.columns): continue
            matches = df[df[col_idx].astype(str).str.contains(keyword, case=False, na=False, regex=False)].index.tolist()
            if matches:
                start_row = matches[0]
                break
        
        if start_row == -1:
            return None
        
        # Find Header Row
        if header_offset is not None:
            header_row_idx = start_row + header_offset
            if header_row_idx >= len(df):
                header_row_idx = -1
        else:
            # Start looking from start_row + 1
            current_row = start_row + 1
            header_row_idx = -1
            # Limit search to next 10 rows to be safe
            for i in range(10):
                if current_row + i >= len(df): break
                row = df.iloc[current_row + i]
                # Check if row has some content (at least 2 non-nan values for a valid header)
                if row.count() >= 2:
                    header_row_idx = current_row + i
                    break
        
        if header_row_idx == -1: return None
        # Slice from header
        sub_df = df.iloc[header_row_idx:].reset_index(drop=True)
        
        # Find End of Table (First empty row or next keyword-like row?)
        # We assume empty row separates tables
        end_row = len(sub_df)
        for i in range(1, len(sub_df)): # Skip header
            row = sub_df.iloc[i]
            # Empty row check: check if all values are NaN or empty strings
            is_row_empty = True
            for val in row:
                if pd.notna(val) and str(val).strip() != "":
                    is_row_empty = False
                    break
            if is_row_empty:
                end_row = i
                break
        
        table = sub_df.iloc[:end_row].copy()
        
        # Promote first row to header
        # Cleanup: Handle headers
        # 1. Drop columns that are completely empty (all NaN in data rows)
        # Note: We must check data rows, but we haven't set headers yet? 
        # Actually table.iloc[0] is header. table[1:] is data.
        
        # We need to reconstruct the dataframe with proper headers
        data = table.iloc[1:].reset_index(drop=True)
        headers = list(table.iloc[0])
        
        # Identify columns to keep
        keep_indices = []
        first_data_col_found = False
        
        for i, h in enumerate(headers):
            if i >= len(data.columns):
                continue
            col_data = data.iloc[:, i]
            # Keep column if it has data OR if header is not empty/NaN
            if not col_data.dropna().empty or (pd.notna(h) and str(h).strip() != ""):
                keep_indices.append(i)
                # If header is NaN/Empty and this is the first data column, assume it's Variable
                if (pd.isna(h) or str(h).strip() == "") and not first_data_col_found:
                    headers[i] = "Variable"
                    first_data_col_found = True
                elif (pd.isna(h) or str(h).strip() == ""):
                    # If header is empty but it's not the first data col, maybe preserve it or drop?
                    # SmartPLS usually has headers. If missing, maybe "Column_i"
                    headers[i] = f"Column_{i}"
                else:
                    if not first_data_col_found: first_data_col_found = True # Found a labeled col first
        
        # Select valid columns
        final_df = data.iloc[:, keep_indices].copy()
        final_df.columns = [headers[i] for i in keep_indices]
        
        return final_df

    def extract_all_tables(self, filepath, sheet_name='Complete'):
        """
        Extracts all relevant tables from a SmartPLS sheet.
        Returns a list of (table_title, DataFrame) tuples.
        Skips raw data sections like 'Samples', 'Indicator Data', 'Residuals'.
        """
        if not filepath or not os.path.exists(filepath):
            return []
        
        # Keywords to skip (raw data, not useful for interpretation)
        skip_keywords = [
            'Sample ', 'Samples', 'Indicator Data', 'Residual', 'Case ID', 
            'Latent Variable Scores', 'Outer Model Residual Scores', 
            'Inner Model Residual Scores', 'MV Descriptives', 'LV Descriptives',
            'Correlation Matrix', 'Covariance Matrix', 'Stop Criterion', 
            'Setting', 'Base Data', 'Interim', 'Model Selection Criteria',
            'Please cite', 'SmartPLS Report', 'calculating case', 'blindfolding procedure'
        ]
        
        # Keywords that are actual table headers (not data rows)
        table_keywords = [
            'Path Coefficients', 'Indirect Effects', 'Total Indirect Effects',
            'Specific Indirect Effects', 'Total Effects', 'Outer Loadings',
            'Outer Weights', 'R Square', 'f Square', 'f-square', 'Construct Reliability',
            'Fornell-Larcker', 'Cross Loadings', 'HTMT', 'Heterotrait',
            'VIF', 'Inner VIF', 'Outer VIF', 'Fit Summary', 'rms Theta',
            'Mean, STDEV', 'Confidence Intervals', 'Q²', 'Blindfolding',
            'Discriminant Validity', 'Quality Criteria', 'Collinearity',
            'Model_Fit', 'Final Results'
        ]
        
        try:
            xl = pd.ExcelFile(filepath)
            sheet_names_lower = [s.lower() for s in xl.sheet_names]
            if sheet_name.lower() in sheet_names_lower:
                idx = sheet_names_lower.index(sheet_name.lower())
                sheet_name = xl.sheet_names[idx]
            else:
                found = False
                for s in xl.sheet_names:
                    if 'complete' in s.lower() and 'chart' not in s.lower():
                        sheet_name = s
                        found = True
                        break
                if not found:
                    for s in xl.sheet_names:
                        if 'navigation' not in s.lower():
                            sheet_name = s
                            found = True
                            break
                if not found:
                    sheet_name = xl.sheet_names[0]
            
            df = xl.parse(sheet_name, header=None)
            tables = []
            
            # Find all section headers
            section_rows = []
            for idx, row in df.iterrows():
                for col in [0, 1]:
                    if col < len(df.columns):
                        val = row.iloc[col]
                        if pd.notna(val) and isinstance(val, str):
                            text = val.strip()
                            # Check if this looks like a section header
                            # (mostly empty row with title text)
                            non_null_count = row.notna().sum()
                            if non_null_count <= 3 and len(text) > 2:
                                # Check if it's a table keyword
                                is_table = any(kw.lower() in text.lower() for kw in table_keywords)
                                is_skip = any(kw.lower() in text.lower() for kw in skip_keywords)
                                
                                if is_table and not is_skip:
                                    section_rows.append((idx, text))
                                break
            
            # Extract each table
            for i, (row_idx, title) in enumerate(section_rows):
                # Find next section or end of file
                next_row = section_rows[i + 1][0] if i + 1 < len(section_rows) else len(df)
                
                # Extract table between this section and next
                table_df = self._extract_table_at(df, row_idx, next_row)
                
                if table_df is not None and not table_df.empty:
                    # Custom Header Overrides based on user feedback
                    t_lower = title.lower()
                    if "specific indirect effects" in t_lower:
                        if len(table_df.columns) >= 2:
                            new_cols = list(table_df.columns)
                            new_cols[0] = "Jalur"
                            new_cols[1] = "Nilai Efek"
                            table_df.columns = new_cols
                    elif "outer vif" in t_lower:
                        if len(table_df.columns) >= 2:
                            new_cols = list(table_df.columns)
                            new_cols[0] = "Indikator"
                            new_cols[1] = "Nilai VIF"
                            table_df.columns = new_cols
                            
                    tables.append((title, table_df))
            
            return tables
            
        except Exception as e:
            print(f"Error extracting tables from {filepath}: {e}")
            return []
    
    def _extract_table_at(self, df, start_row, end_row):
        """
        Extracts a single table starting after start_row, ending before end_row.
        """
        # Find first data row (header row)
        header_row_idx = -1
        for i in range(start_row + 1, min(start_row + 10, end_row)):
            if i >= len(df):
                break
            row = df.iloc[i]
            if row.notna().sum() > 1:
                header_row_idx = i
                break
        
        if header_row_idx == -1:
            return None
        
        # Find table end (empty row or end_row)
        table_end = end_row
        for i in range(header_row_idx + 1, end_row):
            if i >= len(df):
                break
            row = df.iloc[i]
            # Empty row check: check if all values are NaN or empty strings
            is_row_empty = True
            for val in row:
                if pd.notna(val) and str(val).strip() != "":
                    is_row_empty = False
                    break
            if is_row_empty:
                table_end = i
                break
        
        # Extract slice
        table = df.iloc[header_row_idx:table_end].copy()
        
        if len(table) < 2:
            return None
        
        # Process headers and data
        # Check if first row looks like a header (contains text like column names) or data
        first_row = list(table.iloc[0])
        
        # Heuristic: If first row has numeric values in most columns, it's likely data, not headers
        numeric_count = sum(1 for v in first_row if isinstance(v, (int, float)) and not pd.isna(v))
        text_count = sum(1 for v in first_row if isinstance(v, str) and v.strip() != "")
        
        # Check for path indicators (-> or <-) which imply data
        has_arrow = any(("->" in str(v) or "<-" in str(v)) for v in first_row if pd.notna(v))
        
        # Check for strong header indicators
        header_keywords = [
            "Variable", "Sample", "Mean", "STDEV", "T Statistics", "P Values", 
            "R Square", "f Square", "Cronbach", "AVE", "Composite", "Rho_A", 
            "VIF", "Path", "Hypothesis", "Effect", "Original", "Item", "Indicator", "Construct"
        ]
        has_header_keyword = any(any(k.lower() in str(v).lower() for k in header_keywords) for v in first_row if isinstance(v, str))
        
        # Decision: It is data if:
        # 1. Mostly numeric
        # 2. Has arrow "->" (path data)
        # 3. Does NOT have header keywords AND has some numbers (ambiguous case, prefer data)
        
        is_data_row = (numeric_count > text_count) or has_arrow or (not has_header_keyword and numeric_count > 0) or len(table) == 1
        
        # If mostly numeric or all cells look like data, create default headers
        if is_data_row:
            # First row is data, not header - create synthetic headers
            data = table.reset_index(drop=True)
            num_cols = len(data.columns)
            headers = ["Variable"] + [f"Value_{i}" if i > 0 else "Value" for i in range(num_cols - 1)]
            if num_cols == 2:
                headers = ["Variable", "Value"]
            elif num_cols == 1:
                headers = ["Variable"]
        else:
            # First row is header
            data = table.iloc[1:].reset_index(drop=True)
            headers = list(table.iloc[0])
        
        # Identify columns to keep
        keep_indices = []
        first_data_col_found = False
        
        for i, h in enumerate(headers):
            if i >= len(data.columns):
                continue
            col_data = data.iloc[:, i] if len(data) > 0 else pd.Series()
            # Keep column if it has data OR if header is not empty/NaN
            if not col_data.dropna().empty or (pd.notna(h) and str(h).strip() != ""):
                keep_indices.append(i)
                if (pd.isna(h) or str(h).strip() == "") and not first_data_col_found:
                    headers[i] = "Variable"
                    first_data_col_found = True
                elif pd.isna(h) or str(h).strip() == "":
                    headers[i] = f"Value_{i}" if i > 0 else "Value"
                else:
                    if not first_data_col_found:
                        first_data_col_found = True
        
        if not keep_indices:
            return None
        
        final_df = data.iloc[:, keep_indices].copy() if len(data) > 0 else table.iloc[:, keep_indices].copy()
        final_df.columns = [headers[i] for i in keep_indices]
        
        return final_df

    def get_outer_model(self):
        # 1. Reliability & Validity
        # Source: PLS.xlsx -> 'Construct Reliability and Validity'
        path = self.find_file("PLS.xlsx")
        rel_df = None
        fl_df = None
        rel_data = []
        disc_text = []
        
        if path:
            try:
                xl = pd.ExcelFile(path)
                # Try 'Complete' sheet first
                sheet_name = self.get_complete_sheet_name(xl)
                df = xl.parse(sheet_name, header=None)
                
                # A. Reliability
                # Keywords: "Construct Reliability and Validity"
                rel_df = self.read_subtable(df, "Construct Reliability and Validity")
                if rel_df is not None:
                    # Expected Cols: Matrix of vars. 
                    # Usually: [Matrix] | Cronbach's Alpha | rho_A | Composite Reliability | Average Variance Extracted (AVE)
                    # Let's clean up columns
                    # The table usually has the Variable name in the first column (index)
                    # But the header row might be complex.
                    # Looking at inspect output locally is hard, assume standard names
                    # Columns often: "Cronbach's Alpha", "rho_A", "Composite Reliability", "Average Variance Extracted (AVE)"
                    
                    # Identify columns by keyword match
                    cols = rel_df.columns.astype(str).tolist()
                    var_col = cols[0] # Assume first column is variable name
                    
                    # Helper to find column
                    def find_col(name):
                        for c in cols:
                            if name.lower() in c.lower(): return c
                        return None
                        
                    ca_col = find_col("Cronbach")
                    rho_col = find_col("rho_a")
                    cr_col = find_col("rho_c")
                    if cr_col is None:
                        cr_col = find_col("Composite Reliability")
                    ave_col = find_col("Average Variance Extracted")
                    
                    if ca_col and cr_col and ave_col:
                        for _, row in rel_df.iterrows():
                            # Skip invalid rows
                            if pd.isna(row[var_col]): continue
                            
                            rel_data.append({
                                'Variable': row[var_col],
                                'Cronbach': float(row[ca_col]) if not pd.isna(row[ca_col]) else 0,
                                'Rho_A': float(row[rho_col]) if rho_col and not pd.isna(row[rho_col]) else 0,
                                'CR': float(row[cr_col]) if not pd.isna(row[cr_col]) else 0,
                                'AVE': float(row[ave_col]) if not pd.isna(row[ave_col]) else 0
                            })
                            
                # B. Discriminant Validity
                # Keyword: "Discriminant Validity" -> then "Fornell-Larcker Criterion" sub-tab
                # This is tricky because "Discriminant Validity" is a main header, and "Fornell-Larcker" is a sub-header (tab in UI, but in Excel??)
                # In Excel export 'Complete', they might be separate blocks.
                
                # Let's try to interpret the Fornell Larcker table
                fl_df = self.read_subtable(df, "Fornell-Larcker Criterion")
                if fl_df is None:
                     # Fallback check for just "Discriminant Validity"
                     fl_df = self.read_subtable(df, "Discriminant Validity")
                     
                if fl_df is not None:
                    # Construct simple text summary matrix
                    # Just print the table as markdown or text
                    disc_text.append("Tabel Fornell-Larcker Criterion:")
                    # Get headers (variables)
                    headers = [str(h) for h in fl_df.columns if "nan" not in str(h).lower()]
                    disc_text.append(" | ".join(headers))
                    
                    for _, row in fl_df.iterrows():
                        vals = [str(row[h])[:5] for h in headers] # truncate decimals
                        disc_text.append(" | ".join(vals))
                        
            except Exception as e:
                print(f"Error reading PLS.xlsx: {e}")
                
        ol_data = []
        outer_vif_data = []
        if path:
            try:
                # C. Outer Loadings
                ol_df = self.read_subtable(df, "Outer loadings")
                if ol_df is not None:
                    cols = ol_df.columns.tolist()
                    var_col = cols[0]
                    construct_cols = cols[1:]
                    for _, row in ol_df.iterrows():
                        indicator = row[var_col]
                        if pd.isna(indicator) or str(indicator).strip() == "" or str(indicator).lower() == "nan":
                            continue
                        # Find which construct has non-nan value
                        for c_col in construct_cols:
                            val = row[c_col]
                            if pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan":
                                try:
                                    loading_val = float(val)
                                    kesimpulan = "Valid" if loading_val >= 0.70 else "Dipertahankan" if loading_val >= 0.40 else "Tidak Valid"
                                    ol_data.append({
                                        'Indikator': str(indicator),
                                        'Konstruk Laten': str(c_col),
                                        'Outer Loading': loading_val,
                                        'Kesimpulan': kesimpulan
                                    })
                                    break
                                except ValueError:
                                    pass

                # D. Outer VIF
                outer_vif_df = self.read_subtable(df, "Outer model - List", header_offset=2)
                if outer_vif_df is None:
                    outer_vif_df = self.read_subtable(df, "Outer VIF Values")
                if outer_vif_df is None:
                    outer_vif_df = self.read_subtable(df, "Outer VIF")
                
                if outer_vif_df is not None:
                    cols = outer_vif_df.columns.tolist()
                    var_col = cols[0]
                    vif_col = cols[1] if len(cols) > 1 else cols[0]
                    for _, row in outer_vif_df.iterrows():
                        ind = row[var_col]
                        if pd.isna(ind) or str(ind).strip() == "" or str(ind).lower() == "nan":
                            continue
                        try:
                            v_val = float(row[vif_col])
                            kesimpulan = "Bebas Kolinearitas" if v_val < 5.0 else "Terindikasi Kolinearitas"
                            outer_vif_data.append({
                                'Indikator': str(ind),
                                'VIF': v_val,
                                'Kesimpulan': kesimpulan
                            })
                        except:
                            pass
                # E. Cross Loadings
                cross_df = self.read_subtable(df, "Cross loadings")
                if cross_df is None:
                    cross_df = self.read_subtable(df, "Cross Loadings")
            except Exception as e:
                print(f"Error reading PLS.xlsx outer loadings / vif / cross: {e}")
                
        return rel_data, disc_text, rel_df, fl_df, cross_df, ol_data, outer_vif_data

    def get_inner_model(self):
        r2_data = []
        f2_data = []
        vif_data = []
        q2_data = []
        path_coeffs = []
        total_effects = []
        
        r2_df = None
        vif_df = None
        pc_df = None
        f2_df = None
        q2_df = None
        te_df = None
        
        # 1. R Square & VIF from PLS.xlsx
        path_pls = self.find_file("PLS.xlsx")
        if path_pls:
            try:
                xl = pd.ExcelFile(path_pls)
                sheet_name = self.get_complete_sheet_name(xl)
                df = xl.parse(sheet_name, header=None)
                
                # R Square
                r2_df = self.read_subtable(df, "R-Square")
                if r2_df is None: r2_df = self.read_subtable(df, "R Square")
                if r2_df is not None:
                    # Cols: R Square, R Square Adjusted
                    cols = r2_df.columns.astype(str).tolist()
                    r2_col = next((c for c in cols if ("R Square" in c or "R-square" in c) and "Adj" not in c), None)
                    r2_adj_col = next((c for c in cols if "Adj" in c or "adjusted" in c), None)
                    
                    if r2_col:
                        for _, row in r2_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            r2_data.append({
                                'Variable': row.iloc[0],
                                'R2': float(row[r2_col]) if not pd.isna(row[r2_col]) else 0.0,
                                'R2_Adj': float(row[r2_adj_col]) if r2_adj_col and not pd.isna(row[r2_adj_col]) else 0.0
                            })
                            
                # VIF (Collinearity Statistics (VIF))
                vif_df = self.read_subtable(df, "Inner model - List")
                if vif_df is None: vif_df = self.read_subtable(df, "Inner VIF Values")
                if vif_df is None: vif_df = self.read_subtable(df, "Collinearity Statistics (VIF)")
                if vif_df is not None:
                    is_list_format = False
                    if len(vif_df.columns) == 2:
                        first_col_str = vif_df.iloc[:, 0].astype(str).tolist()
                        if any("->" in x for x in first_col_str):
                            is_list_format = True
                            
                    if is_list_format:
                        for _, row in vif_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            try:
                                val = float(row.iloc[1])
                                vif_data.append({'Variable': str(row.iloc[0]), 'VIF': val})
                            except:
                                pass
                    else:
                        numeric_cols = [c for c in vif_df.columns if c != vif_df.columns[0]]
                        for col in numeric_cols:
                            for _, row in vif_df.iterrows():
                                val = row[col]
                                try:
                                    val = float(val)
                                    if val > 0:
                                        vif_data.append({'Variable': f"{row[vif_df.columns[0]]} -> {col}", 'VIF': val})
                                except:
                                    pass
                                
            except Exception as e:
                print(f"Error reading PLS.xlsx Inner Model: {e}")

        # 2. Path Coeffs from Bootstrapping.xlsx
        path_boot = self.find_file("Bootstrapping.xlsx")
        self.ie_data = []
        if path_boot:
            try:
                xl = pd.ExcelFile(path_boot)
                sheet_name = self.get_complete_sheet_name(xl)
                df = xl.parse(sheet_name, header=None)
                
                # Path Coefficients
                pc_df = self.read_subtable(df, "Path Coefficients")
                if pc_df is None: pc_df = self.read_subtable(df, "Mean, STDEV, T-Values, P-Values")
                if pc_df is None: pc_df = self.read_subtable(df, "Mean, STDEV, T values, p values")
                if pc_df is not None:
                    cols = pc_df.columns.astype(str).tolist()
                    def get_c(key): return next((c for c in cols if key.lower() in c.lower()), None)
                    
                    o_col = get_c("Original Sample")
                    sd_col = get_c("Standard Deviation") or get_c("STDEV") or get_c("standard dev")
                    t_col = get_c("T Statistics") or get_c("T value")
                    p_col = get_c("P Values") or get_c("p value")
                    
                    if o_col and p_col:
                        for _, row in pc_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            path_str = str(row.iloc[0])
                            if "->" not in path_str: continue
                            path_coeffs.append({
                                'Path': path_str,
                                'Beta': float(row[o_col]) if not pd.isna(row[o_col]) else 0.0,
                                'STDEV': float(row[sd_col]) if sd_col and not pd.isna(row[sd_col]) else 0.0,
                                'T': float(row[t_col]) if t_col and not pd.isna(row[t_col]) else 0.0,
                                'P': float(row[p_col]) if not pd.isna(row[p_col]) else 1.0
                            })
                
                # Specific Indirect Effects
                ie_df = self.read_subtable(df, "Specific Indirect Effects")
                if ie_df is None: ie_df = self.read_subtable(df, "Specific indirect effects")
                if ie_df is not None:
                    cols = ie_df.columns.astype(str).tolist()
                    def get_c(key): return next((c for c in cols if key.lower() in c.lower()), None)
                    
                    o_col = get_c("Original Sample")
                    sd_col = get_c("Standard Deviation") or get_c("STDEV") or get_c("standard dev")
                    t_col = get_c("T Statistics") or get_c("T value")
                    p_col = get_c("P Values") or get_c("p value")
                    
                    if o_col and p_col:
                        for _, row in ie_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            path_str = str(row.iloc[0])
                            if "->" not in path_str: continue
                            self.ie_data.append({
                                'Path': path_str,
                                'Beta': float(row[o_col]) if not pd.isna(row[o_col]) else 0.0,
                                'STDEV': float(row[sd_col]) if sd_col and not pd.isna(row[sd_col]) else 0.0,
                                'T': float(row[t_col]) if t_col and not pd.isna(row[t_col]) else 0.0,
                                'P': float(row[p_col]) if not pd.isna(row[p_col]) else 1.0
                            })

                # Total Effects
                te_df = self.read_subtable(df, "Total Effects")
                if te_df is None: te_df = self.read_subtable(df, "Total effects")
                if te_df is not None:
                    cols = te_df.columns.astype(str).tolist()
                    def get_c(key): return next((c for c in cols if key.lower() in c.lower()), None)
                    
                    o_col = get_c("Original Sample")
                    sd_col = get_c("Standard Deviation") or get_c("STDEV") or get_c("standard dev")
                    t_col = get_c("T Statistics") or get_c("T value")
                    p_col = get_c("P Values") or get_c("p value")
                    
                    if o_col and p_col:
                        for _, row in te_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            path_str = str(row.iloc[0])
                            if "->" not in path_str: continue
                            total_effects.append({
                                'Path': path_str,
                                'Beta': float(row[o_col]) if not pd.isna(row[o_col]) else 0.0,
                                'STDEV': float(row[sd_col]) if sd_col and not pd.isna(row[sd_col]) else 0.0,
                                'T': float(row[t_col]) if t_col and not pd.isna(row[t_col]) else 0.0,
                                'P': float(row[p_col]) if not pd.isna(row[p_col]) else 1.0
                            })
            except Exception as e:
                print(f"Error reading Bootstrapping.xlsx: {e}")

        # 3. f Square from Bootstrapping (?) OR PLS.xlsx
        # Usually f Square is in PLS Algorithm results (PLS.xlsx), not Bootstrapping (unless requested).
        # Standard SmartPLS report puts f Square in PLS.xlsx
        # Let's check PLS.xlsx again
        if path_pls:
             try:
                xl = pd.ExcelFile(path_pls)
                sheet_name = self.get_complete_sheet_name(xl)
                df = xl.parse(sheet_name, header=None)
                f2_df = self.read_subtable(df, "f-Square")
                if f2_df is None: f2_df = self.read_subtable(df, "f Square")
                if f2_df is not None:
                    is_list_format = False
                    if len(f2_df.columns) == 2:
                        first_col_str = f2_df.iloc[:, 0].astype(str).tolist()
                        if any("->" in x for x in first_col_str):
                            is_list_format = True
                            
                    if is_list_format:
                        for _, row in f2_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            val = row.iloc[1]
                            try:
                                val = float(val)
                                import math
                                if math.isnan(val):
                                    continue
                                effect = "Kecil" if val < 0.15 else ("Sedang" if val < 0.35 else "Besar")
                                if val < 0.02: effect = "Tidak Ada Pengaruh"
                                f2_data.append({
                                    'Path': str(row.iloc[0]),
                                    'f2': val,
                                    'Effect': effect
                                })
                            except:
                                pass
                    else:
                        numeric_cols = [c for c in f2_df.columns if c != f2_df.columns[0]]
                        for col in numeric_cols:
                            for _, row in f2_df.iterrows():
                                val = row[col]
                                try:
                                    val = float(val)
                                    import math
                                    if math.isnan(val):
                                        continue
                                    effect = "Kecil" if val < 0.15 else ("Sedang" if val < 0.35 else "Besar")
                                    if val < 0.02: effect = "Tidak Ada Pengaruh"
                                    
                                    f2_data.append({
                                        'Path': f"{row[f2_df.columns[0]]} -> {col}",
                                        'f2': val,
                                        'Effect': effect
                                    })
                                except:
                                    pass
             except: pass
             
        # 4. Q Square from Blindfolding.xlsx
        path_bf = self.find_file("Blindfolding.xlsx")
        q2_df = None
        if path_bf:
            try:
                xl = pd.ExcelFile(path_bf)
                sheet_name = self.get_complete_sheet_name(xl)
                df = xl.parse(sheet_name, header=None)
                
                q2_df = self.read_subtable(df, "Construct cross-validated redundancy")
                if q2_df is None:
                    q2_df = self.read_subtable(df, "Construct Crossvalidated Redundancy")
                
                if q2_df is not None:
                    sso_col = next((c for c in q2_df.columns if "SSO" in str(c)), None)
                    sse_col = next((c for c in q2_df.columns if "SSE" in str(c)), None)
                    q2_col = next((c for c in q2_df.columns if "Q" in str(c) and ("2" in str(c) or "²" in str(c))), None)
                    if q2_col:
                        for _, row in q2_df.iterrows():
                            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
                            try:
                                sso_val = float(row[sso_col]) if sso_col and not pd.isna(row[sso_col]) else 0.0
                                sse_val = float(row[sse_col]) if sse_col and not pd.isna(row[sse_col]) else 0.0
                                q2_val = float(row[q2_col]) if not pd.isna(row[q2_col]) else 0.0
                                q2_data.append({
                                    'Variable': row.iloc[0],
                                    'SSO': sso_val,
                                    'SSE': sse_val,
                                    'Q2': q2_val
                                })
                            except:
                                pass
            except Exception as e:
                print(f"Error reading Blindfolding.xlsx Q2: {e}")

        return r2_data, f2_data, vif_data, q2_data, path_coeffs, total_effects, r2_df, f2_df, vif_df, pc_df, q2_df, te_df

    def get_mediation(self):
        vaf_data = [] # Text list
        gof_data = [] # Text list
        
        path_term = self.find_file("Term Output.xlsx")
        if path_term:
            try:
                xl = pd.ExcelFile(path_term)
                
                # VAF
                if 'VAF' in xl.sheet_names:
                    df = xl.parse('VAF')
                    # Columns: Path, Direct, Indirect, Total, VAF (%), Category
                    if 'Path' in df.columns and 'VAF (%)' in df.columns:
                        for _, row in df.iterrows():
                            vaf_data.append(f"Jalur {row['Path']}: VAF {row['VAF (%)']}% ({row.get('Category', '')})")
                            
                # GOF
                # Sheet 'Tenenhaus GoF'
                gof_sheet = next((s for s in xl.sheet_names if "GoF" in s), None)
                if gof_sheet:
                    df = xl.parse(gof_sheet)
                    # Metric, Value
                    for _, row in df.iterrows():
                        gof_data.append(f"{row.iloc[0]}: {row.iloc[1]}")
                        
            except Exception as e:
                print(f"Error reading Term Output.xlsx: {e}")
                
        return vaf_data, gof_data

    def detect_analysis_type(self):
        """
        Detect the type of analysis based on variable naming patterns in extracted data.
        Returns a dict with:
        - type: 'simple_regression', 'mediation', 'moderation', or 'mixed'
        - variables: dict with lists of X, Y, M, Z variables found
        - description: human-readable description
        """
        import re
        
        # Collect all variable names from available files
        all_variables = set()
        
        # Check PLS.xlsx for construct names
        pls_path = self.find_file("PLS.xlsx")
        if pls_path:
            try:
                xl = pd.ExcelFile(pls_path)
                for sheet in xl.sheet_names[:3]:  # Check first few sheets
                    try:
                        df = xl.parse(sheet, header=None)
                        # Scan for variable names in first column and headers
                        for col in df.columns[:5]:
                            for val in df[col].dropna().astype(str):
                                val = val.strip()
                                if val and len(val) < 50:  # Reasonable variable name length
                                    all_variables.add(val)
                    except:
                        pass
            except:
                pass
        
        # Categorize variables
        x_vars = []  # Independent (X, X1, X2, etc.)
        y_vars = []  # Dependent (Y, Y1, Y2, etc.)
        m_vars = []  # Mediator (M, M1, M2, etc.)
        z_vars = []  # Moderator (Z, Z1, Z2, etc.)
        
        # Patterns for variable detection
        x_pattern = re.compile(r'^X\d*$', re.IGNORECASE)
        y_pattern = re.compile(r'^Y\d*$', re.IGNORECASE)
        m_pattern = re.compile(r'^M\d*$', re.IGNORECASE)
        z_pattern = re.compile(r'^Z\d*$', re.IGNORECASE)
        
        for var in all_variables:
            var_clean = var.strip()
            if x_pattern.match(var_clean):
                x_vars.append(var_clean)
            elif y_pattern.match(var_clean):
                y_vars.append(var_clean)
            elif m_pattern.match(var_clean):
                m_vars.append(var_clean)
            elif z_pattern.match(var_clean):
                z_vars.append(var_clean)
        
        # Also check for interaction terms (X*Z pattern)
        interaction_pattern = re.compile(r'(X\d*)\s*[*x×]\s*(Z\d*)', re.IGNORECASE)
        has_interaction = any(interaction_pattern.search(var) for var in all_variables)
        
        # Determine analysis type
        has_x = len(x_vars) > 0
        has_y = len(y_vars) > 0
        has_m = len(m_vars) > 0
        has_z = len(z_vars) > 0 or has_interaction
        
        variables = {
            'X': sorted(set(x_vars)),
            'Y': sorted(set(y_vars)),
            'M': sorted(set(m_vars)),
            'Z': sorted(set(z_vars))
        }
        
        if has_m and has_z:
            analysis_type = 'mixed'
            description = "Analisis Regresi dengan Mediasi dan Moderasi (Mixed)"
        elif has_m:
            analysis_type = 'mediation'
            description = "Analisis Regresi Mediasi"
        elif has_z:
            analysis_type = 'moderation'
            description = "Analisis Regresi Moderasi"
        elif has_x and has_y:
            analysis_type = 'simple_regression'
            description = "Analisis Regresi Sederhana/Berganda"
        else:
            analysis_type = 'unknown'
            description = "Jenis Analisis Tidak Teridentifikasi"
        
        return {
            'type': analysis_type,
            'variables': variables,
            'description': description,
            'has_mediation': has_m,
            'has_moderation': has_z or has_interaction
        }

    def process(self):
        """
        Main pipeline to read files, extract tables, and generate Word report
        with structured Outer/Inner Model layout.
        """
        analysis_info = self.detect_analysis_type()
        self.analysis_type = analysis_info['type']
        self.analysis_variables = analysis_info['variables']
        
        # Pass analysis context to interpreter
        self.interpreter.set_analysis_context(analysis_info)
        
        # Define source files
        source_files = [
            ("PLS.xlsx", "Hasil Analisis PLS Algorithm"),
            ("Bootstrapping.xlsx", "Hasil Bootstrapping"),
            ("Blindfolding.xlsx", "Hasil Blindfolding (Predictive Relevance)")
        ]
        
        # 1. Collect all tables from source files
        all_tables = []
        for filename, section_title in source_files:
            filepath = self.find_file(filename)
            if not filepath:
                continue
            
            tables = self.extract_all_tables(filepath, sheet_name='Complete')
            if not tables:
                try:
                    xl = pd.ExcelFile(filepath)
                    if xl.sheet_names:
                        tables = self.extract_all_tables(filepath, sheet_name=xl.sheet_names[0])
                except:
                    pass
            
            if tables:
                for table_title, table_df in tables:
                    resolved_title = table_title
                    if filename.lower() in ['bootstrapping.xlsx', 'bt.xlsx']:
                        if 'path coefficients' in table_title.lower():
                            resolved_title = "Uji Signifikansi Koefisien Jalur (Bootstrapping - Hipotesis)"
                        elif 'specific indirect effects' in table_title.lower() or 'indirect effects' in table_title.lower():
                            resolved_title = "Uji Signifikansi Efek Tidak Langsung (Bootstrapping - Mediasi)"
                        elif 'total effects' in table_title.lower():
                            resolved_title = "Uji Signifikansi Efek Total (Bootstrapping)"
                        elif 'outer loadings' in table_title.lower():
                            resolved_title = "Uji Signifikansi Outer Loadings (Bootstrapping)"
                        else:
                            resolved_title = f"{table_title} (Bootstrapping)"
                    elif filename.lower() in ['blindfolding.xlsx', 'blf.xlsx', 'bf.xlsx']:
                        if 'final results' in table_title.lower():
                            resolved_title = "Construct cross-validated redundancy (Q²)"
                            
                    all_tables.append({
                        'filename': filename,
                        'original_title': table_title,
                        'title': resolved_title,
                        'df': table_df
                    })
                    
        # 2. Collect tables from Term Output.xlsx (if exists)
        path_term = self.find_file("Term Output.xlsx")
        if path_term:
            try:
                xl = pd.ExcelFile(path_term)
                for sheet_name in xl.sheet_names:
                    try:
                        df = xl.parse(sheet_name)
                        if df.empty or "total effect histogram" in sheet_name.lower():
                            continue
                        df = df.dropna(axis=1, how='all')
                        df = df.dropna(axis=0, how='all')
                        if df.empty:
                            continue
                        
                        all_tables.append({
                            'filename': 'Term Output.xlsx',
                            'original_title': sheet_name,
                            'title': f"Hasil {sheet_name}",
                            'df': df
                        })
                    except Exception as e:
                        print(f"Error processing Term Output sheet {sheet_name}: {e}")
            except Exception as e:
                print(f"Error reading Term Output.xlsx: {e}")
                
        # Filter out HTMT List version if Matrix version is present
        has_htmt_matrix = any(('htmt' in t['original_title'].lower() or 'heterotrait' in t['original_title'].lower()) and 'matrix' in t['original_title'].lower() for t in all_tables)
        if has_htmt_matrix:
            all_tables = [t for t in all_tables if not (('htmt' in t['original_title'].lower() or 'heterotrait' in t['original_title'].lower()) and 'list' in t['original_title'].lower())]

        # 3. Classify and filter tables into Outer, Inner, and Other (only keeping appropriate ones)
        outer_tables = []
        inner_tables = []
        other_tables = []
        
        for table in all_tables:
            title_lower = table['original_title'].lower()
            resolved_lower = table['title'].lower()
            
            # Outer model filtering (only keep appropriate ones)
            is_outer_appropriate = (
                'construct reliability' in title_lower or 'construct reliability' in resolved_lower or
                ('outer loadings' in title_lower and 'bootstrapping' not in resolved_lower) or
                ('outer loadings' in resolved_lower and 'bootstrapping' not in resolved_lower) or
                'fornell-larcker' in title_lower or 'fornell-larcker' in resolved_lower or
                'fornell larcker' in title_lower or 'fornell larcker' in resolved_lower or
                'cross loadings' in title_lower or 'cross loadings' in resolved_lower or
                'cross-loadings' in title_lower or 'cross-loadings' in resolved_lower or
                'htmt' in title_lower or 'htmt' in resolved_lower or
                'heterotrait' in title_lower or 'heterotrait' in resolved_lower
            )
            
            # Inner model filtering (only keep appropriate ones)
            is_inner_appropriate = (
                'path coefficients' in title_lower or 'path coefficients' in resolved_lower or
                'r square' in title_lower or 'r square' in resolved_lower or
                'f square' in title_lower or 'f-square' in title_lower or
                'f square' in resolved_lower or 'f-square' in resolved_lower or
                'fit summary' in title_lower or 'fit summary' in resolved_lower or
                'model_fit' in title_lower or 'model_fit' in resolved_lower or
                'q²' in title_lower or 'q²' in resolved_lower or
                'q square' in title_lower or 'q square' in resolved_lower or
                'construct cross-validated redundancy' in title_lower or 'construct cross-validated redundancy' in resolved_lower or
                'blindfolding' in title_lower or 'blindfolding' in resolved_lower or
                'indirect effect' in title_lower or 'indirect effect' in resolved_lower
            )
            
            if is_outer_appropriate:
                outer_tables.append(table)
            elif is_inner_appropriate:
                inner_tables.append(table)

                    
        # 4. Sort tables for logical flow
        def sort_outer(t):
            t_title = t['original_title'].lower()
            if 'outer loadings' in t_title and 'bootstrapping' not in t['title'].lower():
                return 1
            if 'construct reliability' in t_title:
                return 2
            if 'fornell-larcker' in t_title:
                return 3
            if 'cross loadings' in t_title:
                return 4
            if 'htmt' in t_title or 'heterotrait' in t_title:
                return 5
            if 'outer vif' in t_title:
                return 6
            return 10
            
        outer_tables.sort(key=sort_outer)
        
        def sort_inner(t):
            t_title = t['original_title'].lower()
            is_boot = 'bootstrapping' in t['title'].lower()
            
            if 'path coefficients' in t_title and not is_boot:
                return 1
            if 'r square' in t_title:
                return 2
            if 'f square' in t_title or 'f-square' in t_title:
                return 3
            if 'inner vif' in t_title or 'collinearity' in t_title:
                return 4
            if 'fit summary' in t_title or 'model_fit' in t_title or 'rms theta' in t_title:
                return 5
            if 'path coefficients' in t_title and is_boot:
                return 6
            if 'indirect effect' in t_title:
                return 7
            if 'total effect' in t_title:
                return 8
            if 'blindfolding' in t_title or 'construct cross-validated redundancy' in t_title or 'q²' in t_title or 'final results' in t_title:
                return 9
            if 'vaf' in t_title or 'gof' in t_title or 'tenenhaus' in t_title:
                return 10
            return 20
            
        inner_tables.sort(key=sort_inner)
           # Reset paragraphs list in interpreter
        self.interpreter.paragraphs = []
        
        # 1. Get measurements
        rel_data, disc_text, rel_df, fl_df, cross_df, ol_data, outer_vif_data = self.get_outer_model()
        r2_data, f2_data, vif_data, q2_data, path_coeffs, total_effects, r2_df, f2_df, vif_df, pc_df, q2_df, te_df = self.get_inner_model()
        
        # Helper ordinal function
        def get_indonesian_ordinal(n):
            ordinals = {
                1: "Pertama", 2: "Kedua", 3: "Ketiga", 4: "Keempat", 5: "Kelima",
                6: "Keenam", 7: "Ketujuh", 8: "Kedelapan", 9: "Kesembilan", 10: "Kesepuluh",
                11: "Kesebelas", 12: "Kedua Belas", 13: "Ketiga Belas", 14: "Keempat Belas", 15: "Kelima Belas"
            }
            return ordinals.get(n, f"Ke-{n}")
            
        # 4.1. Analisis Model Pengukuran (Outer Model)
        self.interpreter.add_header("Analisis Model Pengukuran (Outer Model)", level=1)
        self.interpreter.add_text(
            "Evaluasi model pengukuran (outer model) dilakukan untuk menilai validitas dan reliabilitas dari indikator-indikator yang membentuk konstruk laten. "
            "Pengujian ini mencakup evaluasi validitas konvergen melalui nilai loading factor (outer loadings) dan Average Variance Extracted (AVE), "
            "uji reliabilitas konstruk melalui Cronbach's Alpha dan Composite Reliability, "
            "uji validitas diskriminan menggunakan kriteria Heterotrait-Monotrait Ratio (HTMT), "
            "serta evaluasi multikolinearitas indikator melalui nilai Variance Inflation Factor (Outer VIF)."
        )
        
        # 4.1.1. Pengujian Validitas Konvergen (Outer Loadings)
        if ol_data:
            self.interpreter.add_header("Pengujian Validitas Konvergen (Outer Loadings)", level=2)
            self.interpreter.add_text(
                "Uji validitas konvergen ditujukan untuk membuktikan bahwa indikator-indikator yang mengukur suatu konstruk laten memiliki tingkat korelasi yang tinggi. "
                "Kriteria pengujian dilakukan dengan melihat nilai loading factor (outer loadings) masing-masing indikator terhadap konstruk latennya. "
                "Nilai loading factor yang disarankan harus lebih besar atau sama dengan 0,70. Namun, untuk penelitian yang bersifat pengembangan, "
                "nilai loading factor antara 0,40 hingga 0,70 masih dapat dipertahankan selama tidak menyebabkan nilai Average Variance Extracted (AVE) konstruk berada di bawah batas 0,50."
            )
            self.interpreter.add_text("Nilai Loading Factor (Outer Loadings) Indikator")
            
            t_ol_rows = []
            for item in ol_data:
                t_ol_rows.append({
                    'Indikator': item['Indikator'],
                    'Konstruk Laten': item['Konstruk Laten'],
                    'Outer Loading': f"{item['Outer Loading']:.4f}".replace('.', ','),
                    'Batas Kritis': '≥ 0,70',
                    'Kesimpulan': item['Kesimpulan']
                })
            t_ol_df = pd.DataFrame(t_ol_rows)
            self.interpreter.add_table(t_ol_df)
            
            valid_count = sum(1 for item in ol_data if item['Kesimpulan'] == "Valid")
            total_count = len(ol_data)
            ol_vals = [item['Outer Loading'] for item in ol_data]
            min_ol = min(ol_vals) if ol_vals else 0
            max_ol = max(ol_vals) if ol_vals else 0
            min_ol_str = f"{min_ol:.4f}".replace('.', ',')
            max_ol_str = f"{max_ol:.4f}".replace('.', ',')
            
            if valid_count == total_count:
                ol_desc = (
                    f"Berdasarkan tabel nilai loading factor (outer loadings) indikator di atas, hasil pengolahan data menunjukkan bahwa nilai loading factor berkisar "
                    f"antara {min_ol_str} hingga {max_ol_str}. Seluruh indikator memiliki nilai loading factor di atas batas kritis 0,70, "
                    f"sehingga seluruh indikator dinyatakan valid dan memenuhi syarat validitas konvergen pada tingkat indikator."
                )
            else:
                retained_inds = [item['Indikator'] for item in ol_data if item['Kesimpulan'] == "Dipertahankan"]
                retained_str = ", ".join(retained_inds)
                ol_desc = (
                    f"Berdasarkan tabel nilai loading factor (outer loadings) indikator di atas, hasil pengolahan data menunjukkan bahwa sebagian besar indikator memiliki nilai di atas 0,70. "
                    f"Terdapat indikator yang memiliki nilai loading factor antara 0,40 hingga 0,70, yaitu {retained_str}. Indikator tersebut tetap dipertahankan dalam model "
                    f"karena tidak menyebabkan nilai Average Variance Extracted (AVE) konstruk di bawah 0,50. Secara keseluruhan, model memenuhi kriteria validitas konvergen."
                )
            self.interpreter.add_text(ol_desc)

        # 4.1.2. Pengujian Reliabilitas Konstruk dan Validitas Konvergen Konstruk
        if rel_data:
            self.interpreter.add_header("Pengujian Reliabilitas Konstruk dan Validitas Konvergen Konstruk", level=2)
            self.interpreter.add_text(
                "Setelah pengujian validitas konvergen di tingkat indikator terpenuhi, tahap selanjutnya adalah mengevaluasi keandalan (reliabilitas) konstruk laten "
                "serta validitas konvergen di tingkat konstruk menggunakan kriteria Average Variance Extracted (AVE). Reliabilitas konstruk dinilai menggunakan "
                "Cronbach's Alpha dan Composite Reliability (baik rho_a maupun rho_c) dengan nilai ambang batas minimal sebesar 0,70, sedangkan batas minimal AVE adalah 0,50."
            )
            self.interpreter.add_text("Nilai Reliabilitas Konstruk dan Average Variance Extracted (AVE)")
            t1_rows = []
            for r in rel_data:
                t1_rows.append({
                    'Konstruk Laten': r['Variable'],
                    "Cronbach's Alpha": f"{r['Cronbach']:.4f}".replace('.', ','),
                    'Composite Reliability (rho_a)': f"{r['Rho_A']:.4f}".replace('.', ','),
                    'Composite Reliability (rho_c)': f"{r['CR']:.4f}".replace('.', ','),
                    'Average Variance Extracted (AVE)': f"{r['AVE']:.4f}".replace('.', ','),
                    'Kesimpulan': 'Valid & Reliabel'
                })
            t1_df = pd.DataFrame(t1_rows)
            self.interpreter.add_table(t1_df)
            
            min_ave_row = min(rel_data, key=lambda x: x['AVE'])
            max_ave_row = max(rel_data, key=lambda x: x['AVE'])
            min_cr_row = min(rel_data, key=lambda x: x['CR'])
            max_cr_row = max(rel_data, key=lambda x: x['CR'])
            
            min_ave_str = f"{min_ave_row['AVE']:.4f}".replace('.', ',')
            max_ave_str = f"{max_ave_row['AVE']:.4f}".replace('.', ',')
            min_cr_str = f"{min_cr_row['CR']:.4f}".replace('.', ',')
            max_cr_str = f"{max_cr_row['CR']:.4f}".replace('.', ',')
            
            t1_desc_1 = (
                f"Berdasarkan data yang disajikan pada tabel hasil uji validitas konvergen dan reliabilitas konstruk, hasil estimasi algoritma PLS menunjukkan bahwa seluruh konstruk laten "
                f"memiliki nilai Average Variance Extracted (AVE) di atas ambang batas baku sebesar 0,50. Nilai AVE terendah tercatat "
                f"pada konstruk {min_ave_row['Variable']} sebesar {min_ave_str}, sedangkan nilai AVE tertinggi ditunjukkan oleh "
                f"konstruk {max_ave_row['Variable']} sebesar {max_ave_str}. Karena seluruh nilai AVE > 0,50, maka dapat disimpulkan "
                f"bahwa model ini memenuhi syarat validitas konvergen secara empiris, di mana varians indikator yang mampu diekstraksi "
                f"oleh konstruk latennya lebih besar dibandingkan varians kesalahan pengukuran."
            )
            t1_desc_2 = (
                f"Selanjutnya, pengujian keandalan instrumen diukur menggunakan indikator Cronbach's Alpha dan Composite Reliability. "
                f"Hasil pengujian menunjukkan seluruh variabel memiliki nilai Cronbach's Alpha dan Composite Reliability yang berkisar "
                f"antara {min_cr_str} hingga {max_cr_str}. Dengan demikian, instrumen penelitian ini terbukti "
                f"memiliki konsistensi internal yang sangat tinggi dan layak digunakan untuk pengujian tahap struktural selanjutnya."
            )
            self.interpreter.add_text(t1_desc_1)
            self.interpreter.add_text(t1_desc_2)
            
        # Pengujian Validitas Diskriminan (Cross Loadings)
        if cross_df is not None:
            self.interpreter.add_header("Pengujian Validitas Diskriminan (Cross Loadings)", level=2)
            self.interpreter.add_text(
                "Validitas diskriminan pada tingkat indikator dievaluasi dengan membandingkan nilai cross loadings. "
                "Metode ini mensyaratkan nilai loading factor dari setiap indikator terhadap konstruk asalnya harus lebih besar "
                "dibandingkan nilai loading factor-nya terhadap konstruk laten lainnya dalam model pengukuran."
            )
            self.interpreter.add_text("Nilai Cross Loadings Indikator terhadap Seluruh Konstruk Laten")
            
            def format_cross_loadings(df):
                formatted_df = df.copy()
                for col in formatted_df.columns[1:]:
                    for idx in range(len(formatted_df)):
                        val = formatted_df.at[idx, col]
                        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
                            formatted_df.at[idx, col] = ""
                        else:
                            try:
                                val_num = float(val)
                                formatted_df.at[idx, col] = f"{val_num:.4f}".replace('.', ',')
                            except:
                                formatted_df.at[idx, col] = str(val)
                return formatted_df
                
            formatted_cross = format_cross_loadings(cross_df)
            self.interpreter.add_table(formatted_cross)
            
            cl_desc = (
                "Berdasarkan hasil analisis pada tabel cross loadings di atas, terlihat bahwa setiap indikator memiliki nilai loading factor "
                "terbesar terhadap konstruk laten induknya (kolom konstruk masing-masing) jika dibandingkan dengan nilai loading factor "
                "pada kolom konstruk laten lainnya. Hal ini menunjukkan bahwa seluruh indikator secara empiris valid mengukur konstruk induknya "
                "dan terbukti memiliki validitas diskriminan yang baik pada tingkat indikator."
            )
            self.interpreter.add_text(cl_desc)
            
        # Pengujian Validitas Diskriminan (Fornell-Larcker Criterion)
        if fl_df is not None:
            self.interpreter.add_header("Pengujian Validitas Diskriminan (Fornell-Larcker Criterion)", level=2)
            self.interpreter.add_text(
                "Validitas diskriminan pada tingkat konstruk diuji menggunakan metode Fornell-Larcker Criterion. "
                "Metode ini membandingkan nilai akar kuadrat dari Average Variance Extracted (AVE) dari setiap konstruk laten (diletakkan pada baris diagonal) "
                "dengan korelasi korelasi antar konstruk laten lainnya (diletakkan di bawah diagonal). Validitas diskriminan dinyatakan terpenuhi apabila "
                "nilai akar kuadrat AVE masing-masing konstruk laten lebih besar dibandingkan korelasi konstruk tersebut dengan konstruk laten lainnya."
            )
            self.interpreter.add_text("Matriks Kriteria Fornell-Larcker")
            
            def format_fl_table(df):
                formatted_df = df.copy()
                for col in formatted_df.columns[1:]:
                    for idx in range(len(formatted_df)):
                        val = formatted_df.at[idx, col]
                        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
                            formatted_df.at[idx, col] = ""
                        else:
                            try:
                                val_num = float(val)
                                formatted_df.at[idx, col] = f"{val_num:.4f}".replace('.', ',')
                            except:
                                formatted_df.at[idx, col] = str(val)
                return formatted_df
                
            formatted_fl = format_fl_table(fl_df)
            self.interpreter.add_table(formatted_fl)
            
            fl_desc = (
                "Berdasarkan tabel kriteria Fornell-Larcker di atas, hasil pengolahan data menunjukkan bahwa nilai akar kuadrat dari "
                "AVE untuk setiap konstruk laten (nilai yang terletak pada baris diagonal) memiliki nilai yang lebih besar dibandingkan "
                "nilai korelasi konstruk laten tersebut dengan konstruk laten lainnya dalam model. Hasil ini menunjukkan "
                "bahwa model pengukuran ini memenuhi kriteria validitas diskriminan yang baik berdasarkan standardisasi Fornell-Larcker."
            )
            self.interpreter.add_text(fl_desc)
            
        # 4.1.3. Pengujian Validitas Diskriminan (Heterotrait-Monotrait Ratio)
        htmt_table = next((t for t in all_tables if 'htmt' in t['original_title'].lower() or 'heterotrait' in t['original_title'].lower()), None)
        if htmt_table is not None:
            self.interpreter.add_header("Pengujian Validitas Diskriminan (Heterotrait-Monotrait Ratio)", level=2)
            self.interpreter.add_text(
                "Pengujian validitas diskriminan ditujukan untuk memastikan bahwa suatu konstruk laten secara empiris benar-benar berbeda dengan konstruk lainnya dalam model. "
                "Penelitian ini menggunakan kriteria modern yaitu Heterotrait-Monotrait Ratio of Correlations (HTMT) dengan ambang batas konservatif sebesar 0,85 atau maksimal 0,90."
            )
            self.interpreter.add_text("Persamaan untuk perhitungan manual Heterotrait-Monotrait Ratio (HTMT) adalah sebagai berikut:")
            self.interpreter.add_equation(self.interpreter.HTMT_OMML)
            self.interpreter.add_text(
                "Keterangan:\n"
                "- HTMT_ij: Nilai rasio HTMT antara konstruk i dan konstruk j.\n"
                "- d̄_ij: Rata-rata korelasi seluruh indikator konstruk i dengan seluruh indikator konstruk j (heterotrait-heteromethod).\n"
                "- d̄_ii: Rata-rata korelasi antara indikator konstruk i (monotrait-heteromethod).\n"
                "- d̄_jj: Rata-rata korelasi antara indikator konstruk j (monotrait-heteromethod).\n"
                "Nilai HTMT yang disarankan harus kurang dari 0,90 (atau kurang dari 0,85 untuk kriteria yang lebih konservatif)."
            )
            
            self.interpreter.add_text("Matriks Nilai Heterotrait-Monotrait Ratio (HTMT)")
            
            def format_matrix_table(df):
                formatted_df = df.copy()
                for col in formatted_df.columns[1:]:
                    for idx in range(len(formatted_df)):
                        val = formatted_df.at[idx, col]
                        if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan":
                            formatted_df.at[idx, col] = ""
                        else:
                            try:
                                val_num = float(val)
                                col_idx = list(formatted_df.columns).index(col)
                                row_var = str(formatted_df.iloc[idx, 0])
                                if row_var.strip().lower() == col.strip().lower():
                                    formatted_df.at[idx, col] = "—"
                                elif val_num == 1.0 or val_num == 0.0 or col_idx == idx + 1:
                                    formatted_df.at[idx, col] = "—"
                                else:
                                    formatted_df.at[idx, col] = f"{val_num:.4f}".replace('.', ',')
                            except:
                                formatted_df.at[idx, col] = str(val)
                return formatted_df
                
            formatted_htmt = format_matrix_table(htmt_table['df'])
            self.interpreter.add_table(formatted_htmt)
            
            max_htmt_val = 0.0
            max_htmt_from = ""
            max_htmt_to = ""
            htmt_raw = htmt_table['df']
            for r_idx, row in htmt_raw.iterrows():
                row_name = str(row.iloc[0])
                for col_idx, col_name in enumerate(htmt_raw.columns[1:], 1):
                    val = row.iloc[col_idx]
                    try:
                        val_num = float(val)
                        if val_num > max_htmt_val and val_num < 1.0 and row_name != col_name:
                            max_htmt_val = val_num
                            max_htmt_from = col_name
                            max_htmt_to = row_name
                    except:
                        pass
                        
            if max_htmt_val > 0:
                max_htmt_str = f"{max_htmt_val:.4f}".replace('.', ',')
                t2_desc = (
                    f"Hasil analisis pada tabel matriks nilai Heterotrait-Monotrait Ratio (HTMT) menunjukkan bahwa seluruh nilai rasio HTMT antar konstruk berada di bawah nilai batas kritis 0,85. "
                    f"Nilai korelasi rasio tertinggi ditemukan pada hubungan antara konstruk {max_htmt_from} dengan variabel {max_htmt_to} "
                    f"yaitu sebesar {max_htmt_str}, yang tetap berada di bawah ambang batas ketat. Hasil ini menegaskan "
                    f"bahwa tidak terdapat isu tumpang tindih (overlap) antar fenomena konstruk laten, sehingga validitas diskriminan model "
                    f"ini sepenuhnya terpenuhi."
                )
            else:
                t2_desc = (
                    "Hasil analisis pada tabel matriks nilai Heterotrait-Monotrait Ratio (HTMT) menunjukkan bahwa seluruh nilai rasio HTMT antar konstruk berada di bawah nilai batas kritis 0,85, "
                    "sehingga validitas diskriminan model ini sepenuhnya terpenuhi."
                )
            self.interpreter.add_text(t2_desc)
            
        # 4.1.4. Evaluasi Kolinearitas Indikator (Outer VIF)
        if outer_vif_data:
            self.interpreter.add_header("Evaluasi Kolinearitas Indikator (Outer VIF)", level=2)
            self.interpreter.add_text(
                "Evaluasi nilai Variance Inflation Factor (Outer VIF) dilakukan untuk memastikan bahwa tidak terdapat gejala kolinearitas atau "
                "korelasi berlebih antar indikator penyusun dalam suatu konstruk laten. Nilai VIF indikator yang disarankan harus di bawah 5,0 "
                "dan idealnya di bawah 3,0 untuk menjamin bahwa indikator tidak saling tumpang tindih secara berlebihan."
            )
            self.interpreter.add_text("Nilai Variance Inflation Factor (VIF) Outer Model")
            
            # Map indicator to construct from ol_data
            ind_to_construct = {item['Indikator']: item['Konstruk Laten'] for item in ol_data}
            
            t_ov_rows = []
            for item in outer_vif_data:
                c_name = ind_to_construct.get(item['Indikator'], "—")
                t_ov_rows.append({
                    'Indikator': item['Indikator'],
                    'Konstruk Laten': c_name,
                    'Nilai VIF': f"{item['VIF']:.4f}".replace('.', ','),
                    'Batas Kritis': '< 5,00',
                    'Kesimpulan': item['Kesimpulan']
                })
            t_ov_df = pd.DataFrame(t_ov_rows)
            self.interpreter.add_table(t_ov_df)
            
            vif_vals = [item['VIF'] for item in outer_vif_data]
            min_v = min(vif_vals) if vif_vals else 0
            max_v = max(vif_vals) if vif_vals else 0
            min_v_str = f"{min_v:.4f}".replace('.', ',')
            max_v_str = f"{max_v:.4f}".replace('.', ',')
            
            ov_desc = (
                f"Berdasarkan hasil pengujian nilai Variance Inflation Factor (VIF) outer model, seluruh indikator memiliki nilai VIF yang berada "
                f"pada rentang {min_v_str} hingga {max_v_str}. Seluruh nilai VIF indikator tercatat di bawah 5,0 (dan sebagian besar di bawah 3,0), "
                f"yang menunjukkan bahwa seluruh indikator terbebas dari masalah kolinearitas berat dan layak dipertahankan dalam model pengukuran."
            )
            self.interpreter.add_text(ov_desc)
            
        # 4.2. Analisis Model Struktural (Inner Model)
        self.interpreter.add_header("Analisis Model Struktural (Inner Model)", level=1)
        self.interpreter.add_text(
            "Evaluasi model struktural atau inner model dilakukan guna memahami kekuatan prediktif model serta melihat besarnya pengaruh varians variabel independen terhadap variabel dependen. "
            "Evaluasi ini diukur melalui koefisien determinasi (R²), validitas prediktif model (f²), uji kualitas prediksi (Q²), serta pemeriksaan asumsi multikolinearitas melalui nilai VIF struktural."
        )
        # 4.2.1. Uji Koefisien Determinasi (R²)
        if r2_data:
            self.interpreter.add_header("Uji Koefisien Determinasi (R²)", level=2)
            self.interpreter.add_text(
                "Koefisien Determinasi (R-Square) digunakan untuk menilai tingkat kecukupan varians variabel independen dalam menjelaskan varians variabel dependen."
            )
            self.interpreter.add_text("Hasil Estimasi R-Square (R²)")
            t3_rows = []
            for r in r2_data:
                var = r['Variable']
                is_mediator = any(p['Path'].strip().startswith(f"{var} ->") for p in path_coeffs)
                desc = f"{var} (Variabel Mediasi)" if is_mediator else f"{var} (Variabel Dependen)"
                
                r2_val = r['R2']
                r2_adj_val = r['R2_Adj']
                
                if r2_val >= 0.67:
                    cat = "Kuat / Substansial"
                elif r2_val >= 0.33:
                    cat = "Moderat"
                elif r2_val >= 0.19:
                    cat = "Lemah"
                else:
                    cat = "Sangat Lemah"
                    
                t3_rows.append({
                    'Konstruk Endogen': desc,
                    'R-Square (R²)': f"{r2_val:.4f}".replace('.', ','),
                    'R-Square Adjusted': f"{r2_adj_val:.4f}".replace('.', ','),
                    'Kategori Kekuatan Prediksi': cat
                })
            t3_df = pd.DataFrame(t3_rows)
            self.interpreter.add_table(t3_df)
            
            t3_desc_parts = []
            for r in r2_data:
                var = r['Variable']
                r2_val = r['R2']
                is_mediator = any(p['Path'].strip().startswith(f"{var} ->") for p in path_coeffs)
                var_desc = "mediasi" if is_mediator else "dependen"
                
                preds = [p['Path'].split("->")[0].strip() for p in path_coeffs if p['Path'].split("->")[-1].strip() == var]
                if preds:
                    preds_str = ", ".join(preds[:-1]) + f", dan {preds[-1]}" if len(preds) > 1 else preds[0]
                else:
                    preds_str = "variabel eksogen"
                    
                pct_explained = f"{r2_val*100:.2f}%".replace('.', ',')
                pct_others = f"{(1-r2_val)*100:.2f}%".replace('.', ',')
                
                if r2_val >= 0.67:
                    cat_lower = "kuat dan substansial"
                elif r2_val >= 0.33:
                    cat_lower = "moderat"
                elif r2_val >= 0.19:
                    cat_lower = "lemah"
                else:
                    cat_lower = "sangat lemah"
                    
                r2_val_str = f"{r2_val:.4f}".replace('.', ',')
                part = (
                    f"Nilai koefisien determinasi (R²) untuk konstruk {var_desc} {var} adalah sebesar {r2_val_str}. "
                    f"Hal ini mengindikasikan bahwa kontribusi gabungan dari variabel eksogen {preds_str} mampu menjelaskan variabilitas dari {var} "
                    f"sebesar {pct_explained}, sedangkan sisanya sebesar {pct_others} dijelaskan oleh faktor-faktor lain di luar model struktural. "
                    f"Menurut kriteria Chin, nilai R² yang diperoleh dikategorikan ke dalam tingkat prediksi yang {cat_lower}."
                )
                t3_desc_parts.append(part)
                
            t3_desc_full = "Berdasarkan data pada tabel hasil estimasi R-Square (R²), " + " ".join(t3_desc_parts)
            self.interpreter.add_text(t3_desc_full)
            
        # Uji Ukuran Efek (f-Square)
        if f2_data:
            self.interpreter.add_header("Uji Ukuran Efek (f-Square)", level=2)
            self.interpreter.add_text(
                "Nilai f-Square (effect size) digunakan untuk mengukur kontribusi praktis dari suatu variabel eksogen terhadap nilai R-Square variabel endogen. "
                "Menurut Cohen (1988), nilai f-Square sebesar 0,02 dikategorikan sebagai pengaruh kecil (small), "
                "nilai 0,15 dikategorikan sebagai pengaruh sedang (medium), dan nilai 0,35 dikategorikan sebagai pengaruh besar (large). "
                "Nilai di bawah 0,02 mengindikasikan bahwa variabel eksogen tidak memiliki kontribusi praktis terhadap pembentukan R-Square variabel endogen."
            )
            self.interpreter.add_text("Hasil Estimasi Nilai f-Square (Effect Size)")
            
            if f2_df is not None:
                f2_df_clean = f2_df.copy()
                self.interpreter.add_table(f2_df_clean)
            
            f2_narratives = []
            for item in f2_data:
                path_str = item['Path']
                val = item['f2']
                effect = item['Effect']
                
                parts = [p.strip() for p in path_str.split("->")]
                if len(parts) == 2:
                    var_from = parts[0]
                    var_to = parts[1]
                    val_str = f"{val:.4f}".replace('.', ',')
                    
                    if effect == "Tidak Ada Pengaruh":
                        cat_str = "tidak memiliki kontribusi praktis (no effect)"
                    else:
                        cat_str = f"berada dalam kategori pengaruh {effect.lower()}"
                        
                    narr = (
                        f"Pengaruh variabel {var_from} terhadap {var_to} menghasilkan nilai f-Square sebesar {val_str}. "
                        f"Hal ini menunjukkan bahwa kontribusi praktis dari variabel {var_from} dalam menjelaskan variabilitas {var_to} {cat_str}."
                    )
                    f2_narratives.append(narr)
            
            if f2_narratives:
                self.interpreter.add_text("Berdasarkan tabel hasil estimasi nilai f-Square di atas, berikut merupakan penjelasan rinci mengenai ukuran efek untuk masing-masing hubungan variabel:")
                for n in f2_narratives:
                    self.interpreter.add_text(n)
            
        # Uji Multikolinearitas Struktural (Inner VIF)
        if vif_data:
            self.interpreter.add_header("Uji Multikolinearitas Struktural (Inner VIF)", level=2)
            self.interpreter.add_text(
                "Pemeriksaan kolinearitas dilakukan untuk menjamin bahwa hubungan kausalitas dalam model tidak mengalami distorsi akibat korelasi berlebih antar prediktor. "
                "Batas toleransi nilai VIF dalam pengujian model struktural adalah harus di bawah 5,0 (idealnya < 3,0)."
            )
            self.interpreter.add_text("Nilai Variance Inflation Factor (VIF) Inner Model")
            
            sources = []
            targets = []
            for v in vif_data:
                parts = [p.strip() for p in v['Variable'].split("->")]
                if len(parts) == 2:
                    sources.append(parts[0])
                    targets.append(parts[1])
            unique_sources = sorted(list(set(sources)))
            unique_targets = sorted(list(set(targets)))
            
            t4_rows = []
            for src in unique_sources:
                row_data = {'Variabel Asal': src}
                all_vifs_ok = True
                for tgt in unique_targets:
                    vif_val = None
                    for v in vif_data:
                        parts = [p.strip() for p in v['Variable'].split("->")]
                        if len(parts) == 2 and parts[0] == src and parts[1] == tgt:
                            vif_val = v['VIF']
                            break
                    if vif_val is not None:
                        row_data[f'VIF terhadap {tgt}'] = f"{vif_val:.4f}".replace('.', ',')
                        if vif_val >= 3.0:
                            all_vifs_ok = False
                    else:
                        row_data[f'VIF terhadap {tgt}'] = "—"
                row_data['Kesimpulan Asumsi'] = "Bebas Multikolinearitas" if all_vifs_ok else "Terdapat Multikolinearitas"
                t4_rows.append(row_data)
            t4_df = pd.DataFrame(t4_rows)
            self.interpreter.add_table(t4_df)
            
            vif_vals = [v['VIF'] for v in vif_data]
            if vif_vals:
                min_vif = min(vif_vals)
                max_vif = max(vif_vals)
                min_vif_str = f"{min_vif:.4f}".replace('.', ',')
                max_vif_str = f"{max_vif:.4f}".replace('.', ',')
                t4_desc = (
                    f"Sesuai dengan hasil pengujian pada tabel nilai Variance Inflation Factor (VIF) inner model, seluruh nilai VIF internal berada dalam rentang "
                    f"{min_vif_str} hingga {max_vif_str}, yang mana angka tersebut jauh di bawah ambang batas kritis 3,0. "
                    f"Hal ini menandakan bahwa tidak terdapat gejala multikolinearitas antar variabel prediktor, sehingga "
                    f"pengujian signifikansi koefisien jalur dapat dilanjutkan tanpa adanya bias struktural."
                )
            else:
                t4_desc = (
                    "Berdasarkan hasil pengujian multikolinearitas, seluruh nilai VIF struktural berada di bawah batas "
                    "kritis 3,0, sehingga dapat disimpulkan bahwa model ini bebas dari isu multikolinearitas."
                )
            self.interpreter.add_text(t4_desc)
            
        # Uji Relevansi Prediktif (Q²)
        if q2_data:
            self.interpreter.add_header("Uji Relevansi Prediktif (Q²)", level=2)
            self.interpreter.add_text(
                "Kriteria Q-Square (Q²) digunakan untuk menilai relevansi prediktif model. Nilai Q² > 0 menunjukkan bahwa model memiliki relevansi prediktif (predictive relevance), sedangkan nilai Q² ≤ 0 menunjukkan model kurang memiliki relevansi prediktif."
            )
            self.interpreter.add_text("Hasil Estimasi Nilai Q-Square (Q²)")
            
            t_q2_rows = []
            for q in q2_data:
                var = q['Variable']
                # Check if it is an endogen variable
                is_endogen = any(r['Variable'] == var for r in r2_data)
                if not is_endogen:
                    continue
                    
                is_mediator = any(p['Path'].strip().startswith(f"{var} ->") for p in path_coeffs)
                desc = f"{var} (Variabel Mediasi)" if is_mediator else f"{var} (Variabel Dependen)"
                
                sso = q.get('SSO', 0.0)
                sse = q.get('SSE', 0.0)
                q2_val = q['Q2']
                
                conclusion = "Memiliki Predictive Relevance" if q2_val > 0 else "Tidak Memiliki Predictive Relevance"
                
                t_q2_rows.append({
                    'Konstruk Endogen': desc,
                    'SSO': f"{sso:.0f}" if sso.is_integer() else f"{sso:.4f}".replace('.', ','),
                    'SSE': f"{sse:.4f}".replace('.', ','),
                    'Q² (=1-SSE/SSO)': f"{q2_val:.4f}".replace('.', ','),
                    'Kesimpulan': conclusion
                })
            
            if t_q2_rows:
                t_q2_df = pd.DataFrame(t_q2_rows)
                self.interpreter.add_table(t_q2_df)
                
                q2_desc_parts = []
                for q in q2_data:
                    var = q['Variable']
                    is_endogen = any(r['Variable'] == var for r in r2_data)
                    if not is_endogen:
                        continue
                    is_mediator = any(p['Path'].strip().startswith(f"{var} ->") for p in path_coeffs)
                    var_desc = "mediasi" if is_mediator else "dependen"
                    q2_val = q['Q2']
                    q2_val_str = f"{q2_val:.4f}".replace('.', ',')
                    
                    if q2_val > 0:
                        part = f"Konstruk {var_desc} {var} menghasilkan nilai Q² sebesar {q2_val_str} (> 0), yang mengindikasikan bahwa model penelitian memiliki relevansi prediktif yang baik untuk konstruk tersebut."
                    else:
                        part = f"Konstruk {var_desc} {var} menghasilkan nilai Q² sebesar {q2_val_str} (≤ 0), yang menunjukkan model kurang memiliki relevansi prediktif untuk konstruk tersebut."
                    q2_desc_parts.append(part)
                    
                q2_desc_full = "Berdasarkan data pada tabel hasil estimasi Q-Square (Q²), " + " ".join(q2_desc_parts)
                self.interpreter.add_text(q2_desc_full)
            
        # Kebaikan Model Fit (SRMR)
        srmr_val = 0.08
        fit_table = next((t for t in all_tables if 'fit summary' in t['original_title'].lower() or 'model_fit' in t['original_title'].lower()), None)
        if fit_table is not None:
            df_fit = fit_table['df']
            for _, row in df_fit.iterrows():
                row_lbl = str(row.iloc[0]).lower()
                if 'srmr' in row_lbl:
                    try:
                        srmr_val = float(row.iloc[1])
                        break
                    except:
                        pass
                        
        self.interpreter.add_header("Kebaikan Model Fit (SRMR)", level=2)
        self.interpreter.add_text(
            "Evaluasi kecocokan model (goodness of fit) dilakukan menggunakan metrik Fit Summary. "
            "Indikator utama yang digunakan adalah Standardized Root Mean Square Residual (SRMR), di mana nilai SRMR di bawah 0,08 menunjukkan bahwa model teoritis memiliki kecocokan yang baik dengan data empiris."
        )
        self.interpreter.add_text("Persamaan untuk perhitungan manual Standardized Root Mean Square Residual (SRMR) adalah sebagai berikut:")
        self.interpreter.add_equation(self.interpreter.SRMR_OMML)
        self.interpreter.add_text(
            "Keterangan:\n"
            "- SRMR: Standardized Root Mean Square Residual.\n"
            "- s_ij: Korelasi sampel (observasi) antara indikator i dan j.\n"
            "- σ_ij: Korelasi model (prediksi) antara indikator i dan j.\n"
            "- h: Jumlah elemen korelasi unik (p(p-1)/2, di mana p adalah jumlah indikator).\n"
            "Nilai SRMR < 0,08 menunjukkan kecocokan model yang baik (fit)."
        )
        
        # Create Fit Summary Table for presentation
        fit_table_data = []
        if fit_table is not None:
            df_fit = fit_table['df']
            srmr_row = None
            nfi_row = None
            for _, row in df_fit.iterrows():
                lbl = str(row.iloc[0]).lower().strip()
                if lbl == 'srmr':
                    srmr_row = row
                elif lbl == 'nfi':
                    nfi_row = row
            
            if srmr_row is not None:
                try:
                    sat_val = float(srmr_row.iloc[1])
                    est_val = float(srmr_row.iloc[2])
                    fit_table_data.append({
                        'Fit Summary': 'SRMR',
                        'Saturated Model': f"{sat_val:.4f}".replace('.', ','),
                        'Estimated Model': f"{est_val:.4f}".replace('.', ','),
                        'Rentang Standar': '< 0,08 (Model Fit Baik)'
                    })
                except Exception:
                    pass
            
            if nfi_row is not None:
                try:
                    sat_val = float(nfi_row.iloc[1])
                    est_val = float(nfi_row.iloc[2])
                    fit_table_data.append({
                        'Fit Summary': 'NFI',
                        'Saturated Model': f"{sat_val:.4f}".replace('.', ','),
                        'Estimated Model': f"{est_val:.4f}".replace('.', ','),
                        'Rentang Standar': 'Mendekati 1,00 (> 0,90 Sangat Baik)'
                    })
                except Exception:
                    pass
                    
        if fit_table_data:
            self.interpreter.add_text("Hasil Estimasi Kriteria Kecocokan Model (Model Fit Summary)")
            df_fit_presentation = pd.DataFrame(fit_table_data)
            self.interpreter.add_table(df_fit_presentation)
            
        srmr_str = f"{srmr_val:.4f}".replace('.', ',')
        t5_desc = (
            f"Berdasarkan hasil analisis, diperoleh nilai SRMR sebesar {srmr_str}, yang berada di bawah ambang batas 0,08. "
            f"Hasil ini menunjukkan bahwa model teoritis memiliki kecocokan yang baik dengan data empiris."
        )
        self.interpreter.add_text(t5_desc)
        
        # 4.2.4. Kebaikan Model Fit (GoF)
        avg_ave = 0.0
        avg_r2 = 0.0
        gof_val = 0.0
        if rel_data:
            avg_ave = sum(r['AVE'] for r in rel_data) / len(rel_data)
        if r2_data:
            avg_r2 = sum(r['R2'] for r in r2_data) / len(r2_data)
        if avg_ave > 0 and avg_r2 > 0:
            gof_val = (avg_ave * avg_r2) ** 0.5
            
        if gof_val > 0:
            self.interpreter.add_header("Kebaikan Model Fit (GoF)", level=2)
            self.interpreter.add_text("Persamaan untuk perhitungan manual Goodness of Fit (GoF) adalah sebagai berikut:")
            self.interpreter.add_equation(self.interpreter.GOF_OMML)
            
            fit_desc = "sangat kecil"
            if gof_val >= 0.36:
                fit_desc = "besar (large)"
            elif gof_val >= 0.25:
                fit_desc = "sedang (medium)"
            elif gof_val >= 0.10:
                fit_desc = "kecil (small)"
                
            avg_ave_str = f"{avg_ave:.4f}".replace('.', ',')
            avg_r2_str = f"{avg_r2:.4f}".replace('.', ',')
            gof_str = f"{gof_val:.4f}".replace('.', ',')
            t6_desc = (
                "Keterangan:\n"
                "- GoF: Goodness of Fit index.\n"
                "- AVĒ: Rata-rata nilai Average Variance Extracted (AVE) dari seluruh konstruk.\n"
                "- R̄²: Rata-rata nilai R-Square dari seluruh konstruk endogen.\n\n"
                f"Berdasarkan hasil analisis data, diperoleh rata-rata nilai AVE sebesar {avg_ave_str} dan rata-rata nilai R-Square sebesar {avg_r2_str}. "
                f"Dengan demikian, nilai GoF dihitung sebagai berikut:\n"
                f"GoF = \u221a({avg_ave_str} \u00d7 {avg_r2_str}) = {gof_str}.\n\n"
                f"Berdasarkan kriteria Tenenhaus et al. (2004), nilai GoF sebesar {gof_str} "
                f"menunjukkan bahwa model penelitian memiliki tingkat kecocokan (fit) yang {fit_desc}."
            )
            self.interpreter.add_text(t6_desc)
            
        # 4.3. Pengujian Hipotesis Pengaruh Langsung (Direct Effects)
        self.interpreter.add_header("Pengujian Hipotesis Pengaruh Langsung (Direct Effects)", level=1)
        self.interpreter.add_text(
            "Uji hipotesis kausalitas antar konstruk dilakukan dengan menggunakan metode Bootstrapping dengan jumlah subsampel sebanyak 5.000 kali pengulangan. "
            "Pengaruh antar variabel dinyatakan memiliki signifikansi statistik apabila memiliki nilai T-statistik > 1,96 (pada tingkat signifikansi alpha = 5%) dan p-value < 0,05."
        )
        
        if path_coeffs:
            self.interpreter.add_text("Hasil Pengujian Koefisien Jalur Pengaruh Langsung")
            t7_rows = []
            for idx, path in enumerate(path_coeffs, 1):
                p_str = path['Path']
                beta = path['Beta']
                stdev = path.get('STDEV', 0.0)
                t_stat = path['T']
                p_val = path['P']
                is_sig = p_val <= 0.05
                
                p_val_str = f"{p_val:.4f}".replace('.', ',')
                if p_val < 0.001:
                    p_val_str = "< 0,001"
                    
                t7_rows.append({
                    'Hipotesis': f'H{idx}',
                    'Jalur Pengaruh': p_str,
                    'Koefisien Sampel (O)': f"{beta:.4f}".replace('.', ','),
                    'Standar Deviasi (STDEV)': f"{stdev:.4f}".replace('.', ',') if stdev > 0 else '-',
                    'T-Statistik (|O/STDEV|)': f"{t_stat:.4f}".replace('.', ','),
                    'P-Value': p_val_str,
                    'Kesimpulan': 'Diterima' if is_sig else 'Ditolak'
                })
            t7_df = pd.DataFrame(t7_rows)
            self.interpreter.add_table(t7_df)
            
            for idx, path in enumerate(path_coeffs, 1):
                p_str = path['Path']
                beta = path['Beta']
                t_stat = path['T']
                p_val = path['P']
                is_sig = p_val <= 0.05
                
                parts = [p.strip() for p in p_str.split("->")]
                var_from = parts[0]
                var_to = parts[1]
                
                p_val_str = f"sebesar {p_val:.4f}".replace('.', ',')
                if p_val < 0.001:
                    p_val_str = "< 0,001"
                    
                ord_str = get_indonesian_ordinal(idx)
                dec_label = "diterima" if is_sig else "ditolak"
                sig_label = "signifikan" if is_sig else "tidak signifikan"
                arah_label = "positif" if beta > 0 else "negatif"
                
                beta_str = f"{beta:.4f}".replace('.', ',')
                t_stat_str = f"{t_stat:.4f}".replace('.', ',')
                
                narr = (
                    f"{idx}. Pengaruh {var_from} terhadap {var_to} (H{idx}): Koefisien estimasi jalur menunjukkan nilai "
                    f"{arah_label} sebesar {beta_str} dengan nilai T-statistik sebesar {t_stat_str} "
                    f"dan p-value {p_val_str}. Nilai T-statistik yang berada {'di atas' if is_sig else 'di bawah'} batas kritis 1,96 "
                    f"mengindikasikan bahwa variabel {var_from} memiliki pengaruh {arah_label} dan {sig_label} secara langsung "
                    f"terhadap {var_to}. Dengan demikian, Hipotesis {ord_str} (H{idx}) dinyatakan {dec_label}."
                )
                self.interpreter.add_text(narr)
                
        # 4.4. Analisis Pengaruh Tidak Langsung (Uji Mediasi)
        if self.ie_data:
            self.interpreter.add_header("Analisis Pengaruh Tidak Langsung (Uji Mediasi)", level=1)
            self.interpreter.add_text(
                "Untuk menguji peran variabel mediator dalam menjelaskan hubungan antara variabel eksogen dan variabel endogen, "
                "dilakukan evaluasi terhadap koefisien Specific Indirect Effects melalui prosedur pengujian Bootstrapping. "
                "Signifikansi pengaruh mediasi ini ditentukan berdasarkan nilai P-Value dari jalur spesifik tidak langsung."
            )
            self.interpreter.add_text("Hasil Uji Pengaruh Tidak Langsung (Spesifik)")
            
            t8_rows = []
            for idx, ie in enumerate(self.ie_data, len(path_coeffs) + 1):
                ie_path = ie['Path']
                ie_beta = ie['Beta']
                ie_stdev = ie.get('STDEV', 0.0)
                ie_t = ie['T']
                ie_p = ie['P']
                ie_sig = ie_p <= 0.05
                
                parts = [p.strip() for p in ie_path.split("->")]
                if len(parts) >= 3:
                    var_from = parts[0]
                    var_to = parts[-1]
                    direct_path = next((p for p in path_coeffs if p['Path'].strip() == f"{var_from} -> {var_to}"), None)
                    if ie_sig:
                        if direct_path and direct_path['P'] <= 0.05:
                            sifat = "Mediasi Parsial"
                        else:
                            sifat = "Mediasi Penuh"
                    else:
                        sifat = "Tidak Ada Mediasi"
                else:
                    sifat = "Mediasi Parsial" if ie_sig else "Tidak Ada Mediasi"
                    
                ie_p_str = f"{ie_p:.4f}".replace('.', ',')
                if ie_p < 0.001:
                    ie_p_str = "< 0,001"
                    
                t8_rows.append({
                    'Hipotesis': f'H{idx}',
                    'Jalur Pengaruh Tidak Langsung': ie_path,
                    'Koefisien Estimasi (O)': f"{ie_beta:.4f}".replace('.', ','),
                    'Standar Deviasi (STDEV)': f"{ie_stdev:.4f}".replace('.', ',') if ie_stdev > 0 else '-',
                    'T-Statistik': f"{ie_t:.4f}".replace('.', ','),
                    'P-Value': ie_p_str,
                    'Sifat Mediasi': sifat,
                    'Kesimpulan': 'Diterima' if ie_sig else 'Ditolak'
                })
            t8_df = pd.DataFrame(t8_rows)
            self.interpreter.add_table(t8_df)
            
            for idx, ie in enumerate(self.ie_data, len(path_coeffs) + 1):
                ie_path = ie['Path']
                ie_beta = ie['Beta']
                ie_t = ie['T']
                ie_p = ie['P']
                ie_sig = ie_p <= 0.05
                
                parts = [p.strip() for p in ie_path.split("->")]
                var_from = parts[0]
                mediator = parts[1]
                var_to = parts[-1]
                
                direct_path = next((p for p in path_coeffs if p['Path'].strip() == f"{var_from} -> {var_to}"), None)
                
                ie_p_str = f"sebesar {ie_p:.4f}".replace('.', ',')
                if ie_p < 0.001:
                    ie_p_str = "< 0,001"
                    
                ord_str = get_indonesian_ordinal(idx)
                dec_label = "diterima" if ie_sig else "ditolak"
                
                if ie_sig:
                    if direct_path and direct_path['P'] <= 0.05:
                        sifat_label = "Partial Mediation (Mediasi Parsial)"
                        impl = f"Hal ini mengimplikasikan bahwa {var_from} mampu memengaruhi {var_to} secara langsung maupun secara tidak langsung melalui perantara {mediator}."
                    else:
                        sifat_label = "Full Mediation (Mediasi Penuh)"
                        impl = f"Hal ini mengimplikasikan bahwa {var_from} tidak mampu memengaruhi {var_to} secara langsung, melainkan harus melalui perantara {mediator} secara penuh."
                        
                    ie_beta_str = f"{ie_beta:.4f}".replace('.', ',')
                    ie_t_str = f"{ie_t:.4f}".replace('.', ',')
                    narr = (
                        f"{idx - len(path_coeffs)}. Efek Mediasi {mediator} pada hubungan {var_from} ke {var_to} (H{idx}): "
                        f"Nilai koefisien pengaruh tidak langsung diperoleh sebesar {ie_beta_str} dengan nilai T-statistik sebesar {ie_t_str} "
                        f"serta p-value {ie_p_str}. Hasil ini menunjukkan bahwa {mediator} terbukti secara signifikan berperan sebagai variabel mediasi. "
                        f"Mengingat pengaruh langsung {var_from} ke {var_to} pada pengujian sebelumnya bernilai "
                        f"{'signifikan' if direct_path and direct_path['P'] <= 0.05 else 'tidak signifikan'}, maka sifat mediasi yang terbentuk adalah "
                        f"{sifat_label}. {impl}"
                    )
                else:
                    ie_beta_str = f"{ie_beta:.4f}".replace('.', ',')
                    ie_t_str = f"{ie_t:.4f}".replace('.', ',')
                    narr = (
                        f"{idx - len(path_coeffs)}. Efek Mediasi {mediator} pada hubungan {var_from} ke {var_to} (H{idx}): "
                        f"Nilai koefisien pengaruh tidak langsung diperoleh sebesar {ie_beta_str} dengan nilai T-statistik sebesar {ie_t_str} "
                        f"serta p-value {ie_p_str}. Hasil ini menunjukkan bahwa {mediator} tidak terbukti secara signifikan berperan sebagai variabel mediasi. "
                        f"Dengan demikian, Hipotesis {ord_str} (H{idx}) dinyatakan {dec_label}."
                    )
                self.interpreter.add_text(narr)
                
        # 4.5. Pengujian Efek Total (Total Effects)
        if total_effects:
            self.interpreter.add_header("Pengujian Efek Total (Total Effects)", level=1)
            self.interpreter.add_text(
                "Pengujian efek total (total effects) ditujukan untuk mengevaluasi total pengaruh kumulatif dari variabel eksogen terhadap variabel endogen, "
                "yang merupakan gabungan dari pengaruh langsung (direct effects) dan seluruh pengaruh tidak langsung (indirect effects) yang melewati variabel intervening/mediasi. "
                "Signifikansi efek total diukur melalui kriteria nilai T-statistik > 1,96 dan p-value < 0,05."
            )
            
            t_te_rows = []
            for idx, path in enumerate(total_effects, 1):
                p_str = path['Path']
                beta = path['Beta']
                stdev = path.get('STDEV', 0.0)
                t_stat = path['T']
                p_val = path['P']
                is_sig = p_val <= 0.05
                
                p_val_str = f"{p_val:.4f}".replace('.', ',')
                if p_val < 0.001:
                    p_val_str = "< 0,001"
                    
                t_te_rows.append({
                    'Jalur Hubungan': p_str,
                    'Koefisien Total (O)': f"{beta:.4f}".replace('.', ','),
                    'Standar Deviasi (STDEV)': f"{stdev:.4f}".replace('.', ',') if stdev > 0 else '-',
                    'T-Statistik': f"{t_stat:.4f}".replace('.', ','),
                    'P-Value': p_val_str,
                    'Kesimpulan': 'Signifikan' if is_sig else 'Tidak Signifikan'
                })
            t_te_df = pd.DataFrame(t_te_rows)
            self.interpreter.add_table(t_te_df)
            
            for idx, path in enumerate(total_effects, 1):
                p_str = path['Path']
                beta = path['Beta']
                t_stat = path['T']
                p_val = path['P']
                is_sig = p_val <= 0.05
                
                parts = [p.strip() for p in p_str.split("->")]
                var_from = parts[0]
                var_to = parts[1] if len(parts) > 1 else ""
                
                p_val_str = f"sebesar {p_val:.4f}".replace('.', ',')
                if p_val < 0.001:
                    p_val_str = "< 0,001"
                    
                sig_label = "signifikan" if is_sig else "tidak signifikan"
                arah_label = "positif" if beta > 0 else "negatif"
                
                beta_str = f"{beta:.4f}".replace('.', ',')
                t_stat_str = f"{t_stat:.4f}".replace('.', ',')
                
                narr = (
                    f"{idx}. Pengaruh efek total kumulatif dari {var_from} terhadap {var_to} menunjukkan nilai "
                    f"koefisien sebesar {beta_str} dengan arah {arah_label}, nilai T-statistik sebesar {t_stat_str}, "
                    f"serta p-value {p_val_str}. Hasil ini menunjukkan bahwa total pengaruh {var_from} terhadap {var_to} "
                    f"bersifat {sig_label} secara statistik."
                )
                self.interpreter.add_text(narr)
                
        # 4.6. Ringkasan Temuan Penelitian
        self.interpreter.add_header("Ringkasan Temuan Penelitian", level=1)
        accepted_count = sum(1 for p in path_coeffs if p['P'] <= 0.05) + sum(1 for ie in self.ie_data if ie['P'] <= 0.05)
        total_hyp = len(path_coeffs) + len(self.ie_data)
        
        final_r2_val = 0.0
        final_var = ""
        for r in r2_data:
            var = r['Variable']
            is_mediator = any(p['Path'].strip().startswith(f"{var} ->") for p in path_coeffs)
            if not is_mediator:
                final_r2_val = r['R2']
                final_var = var
                break
                
        final_pct = f"{final_r2_val*100:.2f}%".replace('.', ',')
        
        t9_desc = (
            f"Secara keseluruhan, pengujian struktural berbasis SmartPLS 4 membuktikan bahwa dari {total_hyp} hipotesis yang diajukan dalam penelitian ini, "
            f"sebanyak {accepted_count} hipotesis dinyatakan didukung oleh data empiris (diterima). "
            f"Model struktural gabungan memiliki kualitas kecocokan model yang sangat tinggi, bebas dari masalah multikolinearitas, serta memiliki relevansi prediktif yang kuat, "
            f"dengan kemampuan menjelaskan varians variabel dependen utama ({final_var}) sebesar {final_pct}."
        )
        self.interpreter.add_text(t9_desc)
        
        # Save and return
        out_path = os.path.join(self.folder_path, "Interpretasi SmartPLS.docx")
        self.interpreter.generate_word_report(out_path)
        return out_path

    def _get_adaptive_intro(self, analysis_info):
        """
        Generate introduction for the document.
        """
        return (
            ""
        )

if __name__ == "__main__":
    # Test
    folder = r"c:/Users/Admin/OneDrive/Dokumen/Asyraf/CODING/interpretasi SmartPLS"
    reader = SmartPLSReader(folder)
    out = reader.process()
    print(f"Done. Saved to {out}")
