
import pandas as pd
import os
import sys
from interpretation import ScientificInterpreter

class SmartPLSReader:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.interpreter = ScientificInterpreter()
        
    def find_file(self, filename):
        path = os.path.join(self.folder_path, filename)
        if os.path.exists(path):
            return path
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
        # Start looking from start_row + 1
        current_row = start_row + 1
        header_row_idx = -1
        
        # Limit search to next 10 rows to be safe
        for i in range(10):
            if current_row + i >= len(df): break
            row = df.iloc[current_row + i]
            # Check if row has some content (more than 1 non-nan value?)
            if row.count() > 1:
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
            # Heuristic: Empty if first 2 cols are NaN or empty strings
            is_empty = True
            for c in range(min(5, len(df.columns))): # Check first 5 cols
                 val = row.iloc[c]
                 if not pd.isna(val) and str(val).strip() != "":
                     is_empty = False
                     break
            if is_empty:
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
            col_data = data.iloc[:, i]
            # Check if column has any non-null data
            if not col_data.dropna().empty:
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
            else:
                 # Column is empty, drop it (unless header is meaningful? usually not)
                 pass
        
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
            'Please cite', 'SmartPLS Report'
        ]
        
        # Keywords that are actual table headers (not data rows)
        table_keywords = [
            'Path Coefficients', 'Indirect Effects', 'Total Indirect Effects',
            'Specific Indirect Effects', 'Total Effects', 'Outer Loadings',
            'Outer Weights', 'R Square', 'f Square', 'Construct Reliability',
            'Fornell-Larcker', 'Cross Loadings', 'HTMT', 'Heterotrait',
            'VIF', 'Inner VIF', 'Outer VIF', 'Fit Summary', 'rms Theta',
            'Mean, STDEV', 'Confidence Intervals', 'Q²', 'Blindfolding',
            'Discriminant Validity', 'Quality Criteria', 'Collinearity',
            'Model_Fit', 'Final Results'
        ]
        
        try:
            xl = pd.ExcelFile(filepath)
            if sheet_name not in xl.sheet_names:
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
            # Empty row check
            non_null = row.notna().sum()
            if non_null <= 1:
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
        
        # Check for path indicators (->) which imply data
        has_arrow = any("->" in str(v) for v in first_row if pd.notna(v))
        
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
            if col_data.empty or not col_data.dropna().empty:
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
                sheet_name = 'Complete' if 'Complete' in xl.sheet_names else xl.sheet_names[0]
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
                    rho_col = find_col("rho_A")
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
                
        return rel_data, disc_text, rel_df, fl_df

    def get_inner_model(self):
        r2_data = []
        f2_data = []
        vif_data = []
        q2_data = []
        path_coeffs = []
        
        r2_df = None
        vif_df = None
        pc_df = None
        f2_df = None
        q2_df = None
        
        # 1. R Square & VIF from PLS.xlsx
        path_pls = self.find_file("PLS.xlsx")
        if path_pls:
            try:
                xl = pd.ExcelFile(path_pls)
                sheet_name = 'Complete' if 'Complete' in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name, header=None)
                
                # R Square
                r2_df = self.read_subtable(df, "R Square")
                if r2_df is not None:
                    # Cols: R Square, R Square Adjusted
                    cols = r2_df.columns.astype(str).tolist()
                    r2_col = next((c for c in cols if "R Square" in c and "Adj" not in c), None)
                    r2_adj_col = next((c for c in cols if "Adj" in c), None)
                    
                    if r2_col:
                        for _, row in r2_df.iterrows():
                            if pd.isna(row[0]): continue
                            r2_data.append({
                                'Variable': row[0],
                                'R2': float(row[r2_col]) if not pd.isna(row[r2_col]) else 0,
                                'R2_Adj': float(row[r2_adj_col]) if r2_adj_col and not pd.isna(row[r2_adj_col]) else 0
                            })
                            
                # VIF (Collinearity Statistics (VIF))
                # Might be split into Inner and Outer? Usually "Inner VIF Values" or similar
                vif_df = self.read_subtable(df, "Inner VIF Values")
                if vif_df is None: vif_df = self.read_subtable(df, "Collinearity Statistics (VIF)")
                # Dealing with matrix format X -> Y ..?
                # Actually VIF in inner model is one value per predictor? No, it's collinearity of predictors.
                # So it's a matrix or list. SmartPLS VIF inner is often a matrix of predictors vs dependent vars.
                # Let's just grab max VIF per variable to simplify?
                # Or just list all high VIFs.
                if vif_df is not None:
                    # It's a matrix usually.
                    # We iterate and find max VIF for each predictor?
                    # Or we just list VIFs > 5
                    # Simple approach: Iterate all numeric cells
                    numeric_cols = [c for c in vif_df.columns if c != vif_df.columns[0]] # Skip first col (Variable names)
                    for col in numeric_cols:
                        for _, row in vif_df.iterrows():
                            val = row[col]
                            try:
                                val = float(val)
                                if val > 0: # valid VIF
                                    # Format: Predictor (row) -> Dependent (col)
                                    # Wait, often rows are Predictors?
                                    vif_data.append({'Variable': f"{row[vif_df.columns[0]]} -> {col}", 'VIF': val})
                            except:
                                pass
                                
            except Exception as e:
                print(f"Error reading PLS.xlsx Inner Model: {e}")

        # 2. Path Coeffs from Bootstrapping.xlsx
        path_boot = self.find_file("Bootstrapping.xlsx")
        if path_boot:
            try:
                xl = pd.ExcelFile(path_boot)
                sheet_name = 'Complete' if 'Complete' in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name, header=None)
                
                # Path Coefficients
                pc_df = self.read_subtable(df, "Path Coefficients") # Or "Mean, STDEV, T-Values, P-Values"
                if pc_df is not None:
                     # Columns: Original Sample (O), Sample Mean (M), Standard Deviation (STDEV), T Statistics, P Values
                     cols = pc_df.columns.astype(str).tolist()
                     def get_c(key): return next((c for c in cols if key.lower() in c.lower()), None)
                     
                     o_col = get_c("Original Sample")
                     t_col = get_c("T Statistics")
                     p_col = get_c("P Values")
                     
                     if o_col and p_col:
                         for _, row in pc_df.iterrows():
                             if pd.isna(row[0]): continue
                             path_coeffs.append({
                                 'Path': row[0], # e.g. X -> Y
                                 'Beta': float(row[o_col]),
                                 'T': float(row[t_col]) if t_col else 0,
                                 'P': float(row[p_col])
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
                sheet_name = 'Complete' if 'Complete' in xl.sheet_names else xl.sheet_names[0]
                df = xl.parse(sheet_name, header=None)
                f2_df = self.read_subtable(df, "f Square")
                if f2_df is not None:
                    # Matrix format usually (Predictor rows, Dependent cols)
                    numeric_cols = [c for c in f2_df.columns if c != f2_df.columns[0]]
                    for col in numeric_cols:
                        for _, row in f2_df.iterrows():
                            val = row[col]
                            try:
                                val = float(val)
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
        if path_bf:
            try:
                xl = pd.ExcelFile(path_bf)
                # Sheet might be "Construct Crossvalidated Redundancy" or "Summary"
                # Let's look for sheet containing "Q" or just parse first sheet
                sheet = next((s for s in xl.sheet_names if "Construct Crossvalidated Redundancy" in s), xl.sheet_names[0])
                df = xl.parse(sheet, header=0) # Assume header 0 for simple files
                
                # Check headers
                # Usually we have SSO, SSE, Q² (=1-SSE/SSO)
                q2_col = next((c for c in df.columns if "Q" in str(c) and ("2" in str(c) or "²" in str(c))), None)
                if q2_col:
                     for _, row in df.iterrows():
                         q2_data.append({
                             'Variable': row[df.columns[0]],
                             'Q2': row[q2_col]
                         })
                     q2_df = df.copy()
                     # Rename first col if needed
                     if q2_df.columns[0] == 0 or str(q2_df.columns[0]).lower() == "nan":
                         q2_df.rename(columns={q2_df.columns[0]: "Variable"}, inplace=True)
            except Exception as e:
                print(f"Error reading Blindfolding.xlsx: {e}")

        return r2_data, f2_data, vif_data, q2_data, path_coeffs, r2_df, f2_df, vif_df, pc_df, q2_df

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
        Process all SmartPLS source files and generate comprehensive interpretation report.
        Extracts ALL relevant tables and provides scientific interpretation for each.
        Adapts interpretations based on detected analysis type (regression, mediation, moderation).
        """
        # Detect analysis type first
        analysis_info = self.detect_analysis_type()
        self.analysis_type = analysis_info['type']
        self.analysis_variables = analysis_info['variables']
        
        # Pass analysis context to interpreter
        self.interpreter.set_analysis_context(analysis_info)
        
        # Define source files and their sections
        source_files = [
            ("PLS.xlsx", "Hasil Analisis PLS Algorithm"),
            ("Bootstrapping.xlsx", "Hasil Bootstrapping"),
            ("Blindfolding.xlsx", "Hasil Blindfolding (Predictive Relevance)"),
            ("Term Output.xlsx", "Output Tambahan")
        ]
        
        # Build adaptive introduction based on analysis type
        intro_text = self._get_adaptive_intro(analysis_info)
        
        section_number = 1
        
        for filename, section_title in source_files:
            filepath = self.find_file(filename)
            
            if not filepath:
                continue
            
            # Extract all tables from this file
            tables = self.extract_all_tables(filepath, sheet_name='Complete')
            
            if not tables:
                # Try first sheet if Complete not found
                try:
                    xl = pd.ExcelFile(filepath)
                    if xl.sheet_names:
                        tables = self.extract_all_tables(filepath, sheet_name=xl.sheet_names[0])
                except:
                    pass
            
            if tables:
                # Add file section header
                self.interpreter.add_section(f"{section_number}. {section_title}", [
                    f"Berikut adalah hasil analisis dari file {filename}:"
                ])
                
                # Add each table with its interpretation
                for table_idx, (table_title, table_df) in enumerate(tables, 1):
                    content = []
                    
                    # Add the table directly (no markdown formatting)
                    content.append(table_df)
                    
                    # Add data-driven interpretation linked to table values
                    interpretation = self.interpreter.get_dynamic_interpretation(table_title, table_df)
                    content.append(interpretation)
                    
                    # Add to document with clean title
                    self.interpreter.add_section(f"{section_number}.{table_idx} {table_title}", content)
                
                section_number += 1
        
        # Also process Term Output.xlsx sheets directly (for VAF, GoF, etc.)
        self._process_term_output()
        
        # Save
        out_path = os.path.join(self.folder_path, "Interpretasi SmartPLS.docx")
        self.interpreter.generate_word_report(out_path)
        return out_path
    
    def _process_term_output(self):
        """
        Process Term Output.xlsx sheets that may have different structure.
        """
        path_term = self.find_file("Term Output.xlsx")
        if not path_term:
            return
        
        try:
            xl = pd.ExcelFile(path_term)
            
            for sheet_name in xl.sheet_names:
                try:
                    df = xl.parse(sheet_name)
                    
                    # Skip empty sheets or unwanted histograms
                    if df.empty or "total effect histogram" in sheet_name.lower():
                        continue
                    
                    # Clean up: Remove all-NaN columns
                    df = df.dropna(axis=1, how='all')
                    df = df.dropna(axis=0, how='all')
                    
                    if df.empty:
                        continue
                    
                    # Add as section
                    content = []
                    content.append(df)
                    
                    # Get data-driven interpretation based on sheet name and data
                    interpretation = self.interpreter.get_dynamic_interpretation(sheet_name, df)
                    content.append(interpretation)
                    
                    self.interpreter.add_section(f"Term Output: {sheet_name}", content)
                    
                except Exception as e:
                    print(f"Error processing Term Output sheet {sheet_name}: {e}")
                    
        except Exception as e:
            print(f"Error reading Term Output.xlsx: {e}")

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
