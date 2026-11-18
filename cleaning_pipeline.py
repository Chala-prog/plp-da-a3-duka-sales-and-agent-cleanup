import pandas as pd
import numpy as np

def clean_agent_data(txns: pd.DataFrame, agents: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Cleans agent transactions and agent network data programmatically.
    
    Parameters:
        txns (pd.DataFrame): Raw agent transactions dataframe.
        agents (pd.DataFrame): Raw agents dataframe.
        
    Returns:
        clean_df (pd.DataFrame): Fully cleaned and merged dataframe with required flag columns.
        log (pd.DataFrame): Audit log recording cleaning steps, dimensions, and affected row counts.
    """
    log_records = []
    
    # Create working copies to prevent SettingWithCopyWarning
    df_txns = txns.copy()
    df_agents = agents.copy()
    
    # 1. Uniqueness: Remove duplicate transaction IDs
    if 'txn_id' in df_txns.columns:
        initial_count = len(df_txns)
        duplicates_count = df_txns.duplicated(subset=['txn_id']).sum()
        df_txns = df_txns.drop_duplicates(subset=['txn_id'], keep='first')
        log_records.append({
            "issue": "Duplicate transaction IDs",
            "dimension": "Uniqueness",
            "column": "txn_id",
            "rows_affected": int(duplicates_count)
        })
        
    # 2. Consistency: Standardize transaction type values
    if 'txn_type' in df_txns.columns:
        df_txns['txn_type'] = df_txns['txn_type'].astype(str).str.strip().str.capitalize()
        df_txns['txn_type'] = df_txns['txn_type'].replace({
            'Dep': 'Deposit',
            'With': 'Withdrawal',
            'Deposits': 'Deposit',
            'Withdrawals': 'Withdrawal'
        })
        
    # 3. Completeness: Handle missing/invalid amounts & create imputation flag
    df_txns['amount_imputed'] = False
    if 'amount' in df_txns.columns:
        df_txns['amount'] = pd.to_numeric(df_txns['amount'], errors='coerce')
        missing_mask = df_txns['amount'].isna()
        missing_count = missing_mask.sum()
        
        if missing_count > 0:
            # Impute missing amounts using median grouped by transaction type
            median_by_type = df_txns.groupby('txn_type')['amount'].transform('median')
            df_txns.loc[missing_mask, 'amount'] = median_by_type[missing_mask]
            # Fallback to overall median if group median is NaN
            overall_median = df_txns['amount'].median()
            df_txns['amount'] = df_txns['amount'].fillna(overall_median)
            
            df_txns.loc[missing_mask, 'amount_imputed'] = True
            log_records.append({
                "issue": "Missing transaction amounts imputed",
                "dimension": "Completeness",
                "column": "amount",
                "rows_affected": int(missing_count)
            })

    # 4. Accuracy: Outlier detection using IQR method
    df_txns['is_outlier'] = False
    if 'amount' in df_txns.columns and len(df_txns) > 0:
        Q1 = df_txns['amount'].quantile(0.25)
        Q3 = df_txns['amount'].quantile(0.75)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        
        outlier_mask = (df_txns['amount'] < lower_bound) | (df_txns['amount'] > upper_bound)
        df_txns.loc[outlier_mask, 'is_outlier'] = True
        
        log_records.append({
            "issue": "Statistical amount outliers detected (IQR method)",
            "dimension": "Accuracy",
            "column": "amount",
            "rows_affected": int(outlier_mask.sum())
        })

    # 5. Validity: Format transaction dates to ISO format
    if 'txn_date' in df_txns.columns:
        df_txns['txn_date'] = pd.to_datetime(df_txns['txn_date'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')

    # 6. Consistency/Validity: Merge with agents data and track unknown agents
    if 'agent_id' in df_txns.columns and 'agent_id' in df_agents.columns:
        clean_df = pd.merge(df_txns, df_agents, on='agent_id', how='left', suffixes=('', '_agent'))
        
        unknown_mask = clean_df['county'].isna() if 'county' in clean_df.columns else pd.Series([False]*len(clean_df))
        unknown_count = unknown_mask.sum()
        
        if 'county' in clean_df.columns:
            clean_df['county'] = clean_df['county'].fillna('Unknown')
            
        log_records.append({
            "issue": "Transactions linked to unknown or missing agents",
            "dimension": "Consistency",
            "column": "agent_id",
            "rows_affected": int(unknown_count)
        })
    else:
        clean_df = df_txns

    # Build final audit log DataFrame
    log = pd.DataFrame(log_records)
    
    return clean_df, log
