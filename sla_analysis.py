import pandas as pd
import numpy as np
from datetime import datetime
from datetime import timedelta
import os
from io import BytesIO

def run_analysis(
    df,
    sla_should_date=None,
    cut_off=None
):  
    # Client side SLA requirements
    sla_config = {
        "AE": {
            "start_col": "SLA关配交接时间",
            "end_col": "首次派送时间",
            "hours_CA": 48,
            "days_CA": 2,
            "hours_nonCA": 96,
            "days_nonCA": 4,
            "target_rate": 0.95,
            "mode": "broad", # 广义妥投
            "total_count": len(df['客户'] == 'AE')
        },
        "CBO": {
            "start_col": "首分拨首次入库时间",
            "end_col": "首次派送时间",
            "hours_CA": 72,
            "days_CA": 3,
            "hours_nonCA": 96,
            "days_nonCA": 4,
            "target_rate": 0.95,
            "mode": "broad",
            "total_count": len(df['客户'] == 'CBO')
        },
        "FBT": {
            "start_col": "首分拨首次入库时间",
            "end_col": "签收成功时间",
            "hours": 48,
            "days": 2, 
            "target_rate": 0.95,
            "mode": "narrow",  # 狭义妥投
            "total_count": len(df['客户'] == 'FBT')
        },
        "CBT": {
            "start_col": "首分拨首次入库时间",
            "end_col": "首次派送时间",
            "hours": 72,
            "days": 3, 
            "target_rate": 0.95,
            "mode": "narrow",
            "total_count": len(df['客户'] == 'CBT')
        },
        "SKA2": {
            "start_col": "SLA关配交接时间",
            "end_col": "签收成功时间",
            "hours": 120,
            "days": 5, 
            "target_rate": 0.98,
            "mode": "narrow",
            "total_count": len(df['客户'] == 'SKA2')
        },
        "TE": {
            "start_col": "首分拨首次入库时间",
            "end_col": "签收成功时间",
            "hours_z12": 72,
            "days_z12": 3, 
            "hours_z34": 120,
            "days_z34": 5, 
            "target_rate": 0.95,
            "mode": "narrow",
            "total_count": len(df['客户'] == 'TE')
        },
        "YW": {
            "start_col": "首分拨首次入库时间",
            "end_col": "签收成功时间",
            "hours": 84,
            "days": 3, 
            "target_rate": 0.95,
            "mode": "narrow",
            "total_count": len(df['客户'] == 'YW')
        },
        "HTE": {
            "start_col": "首分拨首次入库时间",
            "end_col": "首次派送时间",
            "hours_CA": 72,
            "days_CA": 3,
            "hours_nonCA": 120,
            "days_nonCA": 5,
            "target_rate": 0.95,
            "mode": "broad", # 广义妥投
            "total_count": len(df['客户'] == 'HTE')
        },
        "WHUS": {
            "start_col": "首分拨首次入库时间",
            "end_col": "签收成功时间",
            "hours": 120,
            "days": 5,
            "target_rate": 0.96,
            "mode": "narrow", # 狭义妥投
            "total_count": len(df['客户'] == 'WHUS')
        },
        "WHUS-4PX": {
            "start_col": "首分拨首次入库时间",
            "end_col": "签收成功时间",
            "hours": 120,
            "days": 5,
            "target_rate": 0.96,
            "mode": "narrow", # 狭义妥投
            "total_count": len(df['客户'] == 'WHUS-4PX')
        }
    }
    
    # CA order or not
    def not_CA_order(row) -> bool:
        return row["集配站"] in ["HUB_LAX_LAS", "HUB_LAX_PHX"]

    # Zone 1 & 2 or not
    def not_z12_order(row) -> bool:
        return row["集配站"] in ["HUB_LAX_LAS", "HUB_LAX_PHX", "HUB_LAX_BAK", "HUB_LAX_SAC", "HUB_LAX_SFO", "HUB_LAX_UIC"]
    
    # SLA total hours by client and area
    def get_sla_limit_hours(row):
        client = row["客户"]
        cfg = sla_config.get(client)
        if cfg is None:
            return np.nan
        
        if "hours_CA" in cfg:
            if not_CA_order(row):
                return cfg["hours_nonCA"]
            else:
                return cfg["hours_CA"]
        elif "hours_z12" in cfg:
            if not_z12_order(row):
                return cfg["hours_z12"]
            else:
                return cfg["hours_z34"]
                
        return cfg.get("hours", np.nan)
    
    # SLA total days round down
    def get_sla_limit_days(row):
        client = row["客户"]
        cfg = sla_config.get(client)
        if cfg is None:
            return np.nan
        
        if "days_CA" in cfg:
            if not_CA_order(row):
                return cfg["days_nonCA"]
            else:
                return cfg["days_CA"]
        elif "days_z12" in cfg:
            if not_z12_order(row):
                return cfg["days_z12"]
            else:
                return cfg["days_z34"]
        
        return cfg.get("days", np.nan)
    
    # Calculate time difference
    def hours_diff_values(end_time, start_time):
        return (end_time - start_time).total_seconds() / 3600
    
    # SLA start and end time by client
    def get_sla_start_end(row):
        client = row["客户"]
        cfg = sla_config.get(client)
        if cfg is None:
            return (np.nan, np.nan)
        
        start_col = cfg["start_col"]
        end_col   = cfg["end_col"]
        return (row.get(start_col), row.get(end_col))
    
    # Fulfill SLA or not
    def calc_sla_row(row):
        sla_hours = get_sla_limit_hours(row)
        start_time, end_time = get_sla_start_end(row)
    
        if pd.isna(start_time) or np.isnan(sla_hours):
            due_time = pd.NaT
        else:
            if row["客户"] in ["FBT", "CBT", "AE", "HTE", "WHUS", "WHUS-4PX"]:
                due_time = (start_time + timedelta(hours=float(sla_hours))).normalize() + pd.Timedelta(hours=23, minutes=59, seconds=59)
            else:
                due_time = start_time + timedelta(hours=float(sla_hours))
        
        if pd.isna(start_time) or pd.isna(end_time):
            return pd.Series({
                "SLA标准小时": sla_hours,
                "SLA截止时间": due_time,
                "SLA实际小时": np.nan,
                "SLA是否达标": False
            })
        
        duration_hours = (end_time - start_time).total_seconds() / 3600
        is_ok = end_time <= due_time
        
        return pd.Series({
            "SLA标准小时": sla_hours,
            "SLA截止时间": due_time,
            "SLA实际小时": duration_hours,
            "SLA是否达标": is_ok
        })
    
    # Define cols needed
    columns_needed = [
        '面单号',
        '客户',
        '原集配站',
        '集配站.1',
        '原配送站',
        '配送站.1',
        '段码',
        '收件人邮编',
        '分拨大包号',
        '关配交接时间',
        '首分拨首次入库时间',
        '首分拨首次自动分拣时间',
        '首分拨首次人工分拣时间',
        '首分拨首次出库时间',
        '配送站首次入库时间',
        '司机首次领件时间',
        '首次派送时间',
        '派送司机',
        '最新签收失败原因',
        '异常释放时间',
        '配送站归班时间',
        '签收成功时间',
        '末端异常提报时间', 
        '是否错分'
    ]
    
    # Filter and raname
    columns_existing = [c for c in columns_needed if c in df.columns]
    df = df[columns_existing].rename(columns={
        '集配站.1': '集配站',
        '配送站.1': '配送站'
    })
    
    # Unify time cols format
    time_cols=[
        '关配交接时间',
        '首分拨首次入库时间',
        '首分拨首次自动分拣时间',
        '首分拨首次人工分拣时间',
        '首分拨首次出库时间',
        '配送站首次入库时间',
        '司机首次领件时间',
        '首次派送时间',
        '异常释放时间',
        '配送站归班时间',
        '签收成功时间',
        '末端异常提报时间'
    ]
    for i in time_cols:
        if i in df.columns:
            df[i] = pd.to_datetime(df[i], errors='coerce')
    
    # Update AE SLA start time for customhouse time after 9pm
    df["SLA关配交接时间"] = df["关配交接时间"]
    
    mask_late = (df["客户"] == "AE") & df["关配交接时间"].notna() & (df["关配交接时间"].dt.hour >= 21)
    df.loc[mask_late, "SLA关配交接时间"] = (
        df.loc[mask_late, "关配交接时间"].dt.normalize() + pd.Timedelta(days=1)
    )
    
    # Filter only the orders for specified SLA period
    sla_result = df.apply(calc_sla_row, axis=1)
    df = pd.concat([df, sla_result], axis=1)
    df["SLA截止时间"] = pd.to_datetime(df["SLA截止时间"], errors='coerce')
    
    if isinstance(sla_should_date, tuple):
        sla_start, sla_end = sla_should_date # Time period
    else:
        sla_start, sla_end = None, sla_should_date # Single time
    
    mask = pd.Series(True, index=df.index)

    if sla_start is not None and sla_end is not None:
        mask &= df["SLA截止时间"].isna() | df["SLA截止时间"].between(sla_start, sla_end, inclusive="both")
    
    df = df[mask].copy()
    
    # Update total_count of SLA should orders
    client_counts = df["客户"].value_counts()
    
    for client, cnt in client_counts.items():
        if client in sla_config:
            sla_config[client]["total_count"] = int(cnt)
    
    # Merge sorting time
    pos = df.columns.get_loc('首分拨首次自动分拣时间')
    df['首分拨首次分拣时间'] = df[['首分拨首次人工分拣时间', '首分拨首次自动分拣时间']].max(axis=1)
    df = df.drop(columns=[
        '首分拨首次自动分拣时间',
        '首分拨首次人工分拣时间'
    ])
    df.insert(pos, '首分拨首次分拣时间', df.pop('首分拨首次分拣时间'))
    
    client_sla_summary = {}
    
    for client, group in df.groupby("客户"):
        total = len(group)
        fail = (group["SLA是否达标"] == False).sum()
        ok = total - fail
        rate = ok / total
        target = sla_config[client]["target_rate"]
        if client in ["FBT", "SKA2", "YW", "TE", "WHUS", "WHUS-4PX"]:
            print(f"{client}: 总单量 {total}，狭义不达 {total - ok} 单，不达率 {(1-rate)*100:.2f}%")
        else:
            print(f"{client}: 总单量 {total}，广义不达 {total - ok} 单，不达率 {(1-rate)*100:.2f}%")
        
        client_sla_summary[client] = {
            "total": total,
            "ok": ok,
            "fail": fail,
            "rate": rate,
            "target": target,
            "meet_target": rate >= target
        }
    
    fail_clients = [c for c, v in client_sla_summary.items() if not v["meet_target"]]
    fail_df = df[(df["SLA是否达标"] == False)].copy()
    
    # Calculate time consumed in each step
    def hours_diff(end_col, start_col):
        return (df[end_col] - df[start_col]).dt.total_seconds() / 3600
    
    fail_df["耗时_关配→分拨入库"]   = hours_diff("首分拨首次入库时间", "关配交接时间")
    fail_df["耗时_分拨入库→分拨出库"]    = hours_diff("首分拨首次出库时间", "首分拨首次入库时间")
    fail_df["耗时_分拨出库→配送站入库"] = hours_diff("配送站首次入库时间", "首分拨首次出库时间")
    fail_df["耗时_分拨出库→异常登记"] = hours_diff("末端异常提报时间", "首分拨首次出库时间")
    fail_df["耗时_配送站入库→司机领件"] = hours_diff("司机首次领件时间", "配送站首次入库时间")
    fail_df["耗时_司机领件→首次派送"]   = hours_diff("首次派送时间", "司机首次领件时间")
    fail_df["耗时_司机领件→签收成功"]   = hours_diff("签收成功时间", "司机首次领件时间")
    
    # Function for reason analysis
    def pickup_or_not(row):
        if pd.isna(row["耗时_配送站入库→司机领件"]):
            if hours_diff_values(cut_off, row["配送站首次入库时间"]) > 96:
                return ("飘件/包裹在配送站丢失（须严查原因）", "配送")
            else:
                return("DSP严重压单（需警告）/飘件", "配送")
        elif (row["配送站"] == "HUB_LAX_COM" and row["耗时_配送站入库→司机领件"] > 12):
            if row["耗时_配送站入库→司机领件"] >36:
                return("DSP严重压单（需警告）/飘件", "配送")
            else:
                return("DSP压单/飘件", "配送")
        elif row["耗时_配送站入库→司机领件"] > 8:
            if row["耗时_配送站入库→司机领件"] >32:
                return("DSP严重压单（需警告）/飘件", "配送")
            else:
                return("DSP压单/飘件", "配送")
        else:
            if pd.isna(row["耗时_司机领件→首次派送"]):
                if hours_diff_values(cut_off, row["司机首次领件时间"]) > 48:
                    return ("DSP领件两日内未投递（需警告）", "配送")
                else:
                    return("DSP领件未及时投递", "配送")
            elif row["耗时_司机领件→首次派送"] > 16:
                if row["耗时_司机领件→首次派送"] > 42:
                    return ("DSP领件两日内未投递（需警告）", "配送")
                else:
                    return("DSP领件未及时投递", "配送")
            else:
                if row["客户"] in ["FBT", "SKA2", "YW", "TE", "WHUS", "WHUS-4PX"]:
                    if pd.isna(row["签收成功时间"]):
                        return ("DSP因某些原因未投递成功（需确认）", "配送")
                    elif row["耗时_司机领件→签收成功"] > 16:
                        return ("DSP达成投递要求但超时", "配送")
                    else:
                        return ("DSP达成投递要求但略微超时，存在潜在问题（需确认）", "待确认")
                else:
                    return ("DSP达成投递要求但略微超时，存在潜在问题（需确认）", "待确认")
    
    def missort_or_not(row):
        if row["是否错分"] == "是":
            return ("错分", "分拨")
        else:
            return("卡车迟到/分拨数据出库实物未装车/大包不准确或丢失/大包漏扫/错分未打标/系统bug（需确认）", "待确认")
    
    def sort_or_not(row):
        if pd.isna(row["首分拨首次分拣时间"]):
            return ("分拨未分拣", "分拨")
        else:
            if pd.isna(row["首分拨首次出库时间"]):
                return ("分拨未及时出库", "分拨")
            else:
                limit_hours = 12 + 24 * (get_sla_limit_days(row) - 2)
                if row["耗时_分拨入库→分拨出库"] > limit_hours:
                    return ("分拨未及时出库", "分拨")
                else:
                    if pd.isna(row["耗时_分拨出库→配送站入库"]):
                        if pd.isna(row["末端异常提报时间"]):
                            if hours_diff_values(cut_off, row["首分拨首次出库时间"]) > 96:
                                return ("仓配交接差异", "分拨")
                            else:
                                return missort_or_not(row)
                        else:
                            if row["耗时_分拨出库→异常登记"] > 12:
                                if row["耗时_分拨出库→异常登记"] > 96:
                                    return ("仓配交接差异", "分拨")
                                else:
                                    return missort_or_not(row)
                            else:
                                return pickup_or_not(row)
                    elif row["耗时_分拨出库→配送站入库"] > 0.25:
                        if row["集配站"] == "HUB_LAX_COM":
                            return missort_or_not(row)
                        else:
                            if row["耗时_分拨出库→配送站入库"] > 12:
                                return missort_or_not(row)
                            else:
                                return pickup_or_not(row)
                    else:
                        return pickup_or_not(row)
    
    # Run the reason analysis through df
    results = []
    
    for idx, row in fail_df.iterrows():
        # if row["客户"] not in fail_clients:
        #     reason, duty = ("客户时效SLA达标不追究问题", "无")
        # else:
        if pd.isna(row["首分拨首次入库时间"]):
            reason, duty = ("分拨未入库", "分拨")
        else:
            if row["客户"] in ["AE", "SKA2"]:
                if row["耗时_关配→分拨入库"] > 24:
                    reason, duty = ("分拨入库过晚", "分拨")
                else:
                    reason, duty = sort_or_not(row)
            else:
                reason, duty = sort_or_not(row)
    
        results.append({
            "index": idx,
            "链路问题归因": reason,
            "主要责任方": duty
        })
    
    result_df = pd.DataFrame(results)
    
    # Merge result into df
    fail_df = fail_df.merge(result_df, left_index=True, right_on="index", how="left").drop(columns=["index"])
    
    # ===== 1. Summary for all orders together =====
    total_orders = len(df)
    total_fail = len(fail_df)
    overall_fail_rate = total_fail / total_orders if total_orders > 0 else np.nan
    
    summary_all = (
        fail_df
        .groupby(["链路问题归因", "主要责任方"])
        .size()
        .reset_index(name="问题单量")
    )
    
    summary_all["占比_numeric"] = summary_all["问题单量"] / total_orders
    summary_all = summary_all.sort_values("占比_numeric", ascending=False)  # Sort from high to low
    summary_all["占整体总单量比"] = (summary_all["占比_numeric"] * 100).round(2).astype(str) + "%"
    summary_all = summary_all.drop(columns=["占比_numeric"])
    
    overall_info = pd.DataFrame([
        ["总单量", total_orders],
        ["未达标单量", total_fail],
        ["未达标率", f"{overall_fail_rate*100:.2f}%" if total_orders > 0 else ""]
    ], columns=["指标", "值"])
    
    # ===== 2. Summary by client =====
    client_total = (
        df["客户"]
        .value_counts()
        .rename_axis("客户")
        .reset_index(name="客户总单量")
    )
    
    summary_by_client = (
        fail_df
        .groupby(["客户", "链路问题归因", "主要责任方"])
        .size()
        .reset_index(name="问题单量")
    )
    
    summary_by_client = summary_by_client.merge(client_total, on="客户", how="left")
    summary_by_client["占比_numeric"] = summary_by_client["问题单量"] / summary_by_client["客户总单量"]
    summary_by_client = summary_by_client.sort_values(
        ["客户", "占比_numeric"],
        ascending=[True, False]
    )
    summary_by_client["占客户总单量比"] = (
        summary_by_client["占比_numeric"] * 100
    ).round(2).astype(str) + "%"
    summary_by_client = summary_by_client.drop(columns=["占比_numeric"])
    
    client_sla_summary = {}
    
    for client, group in df.groupby("客户"):
        total_c = len(group)
        fail_c  = (group["SLA是否达标"] == False).sum()
        ok_c    = total_c - fail_c
    
        success_rate = ok_c / total_c if total_c > 0 else np.nan
    
        cfg = sla_config.get(client, {})
        target = cfg.get("target_rate", np.nan)
    
        meet = (not np.isnan(target)) and (success_rate >= target)
    
        client_sla_summary[client] = {
            "total": total_c,
            "fail": fail_c,
            "ok": ok_c,
            "success_rate": success_rate,
            "fail_rate": fail_c / total_c if total_c > 0 else np.nan,
            "target": target,
            "meet_target": meet,
        }
    
    # ===== 3. Summary by hub =====
    hub_total = (
        df["集配站"]
        .value_counts()
        .rename_axis("集配站")
        .reset_index(name="站点总单量")
    )
    
    summary_by_hub = (
        fail_df
        .groupby(["集配站", "链路问题归因", "主要责任方"])
        .size()
        .reset_index(name="问题单量")
    )
    
    summary_by_hub = summary_by_hub.merge(hub_total, on="集配站", how="left")
    summary_by_hub["占比_numeric"] = summary_by_hub["问题单量"] / summary_by_hub["站点总单量"]
    summary_by_hub = summary_by_hub.sort_values(
        ["集配站", "占比_numeric"],
        ascending=[True, False]
    )
    summary_by_hub["占总单量比"] = (
        summary_by_hub["占比_numeric"] * 100
    ).round(2).astype(str) + "%"
    summary_by_hub = summary_by_hub.drop(columns=["占比_numeric"])
    
    # === 汇总到 集配站 级别（消除 duplicate） ===
    hub_overall = (
        summary_by_hub
        .groupby("集配站", as_index=False)
        .agg(
            站点总单量=("站点总单量", "first"),   
            问题单量=("问题单量", "sum")      
        )
    )
    
    hub_overall["占集配站总单量比"] = (
        hub_overall["问题单量"] / hub_overall["站点总单量"] * 100
    ).round(2).astype(str) + "%"
    
    hub_sla_summary = {}
    hub_sta_summary = {}
    
    for hub, group in df.groupby("集配站"):
        total_h = len(group)
        fail_h  = (group["SLA是否达标"] == False).sum()
        ok_h    = total_h - fail_h
    
        success_rate = ok_h / total_h if total_h > 0 else np.nan
    
        cfg = sla_config.get(hub, {})
        target = cfg.get("target_rate", np.nan)
    
        hub_sla_summary[hub] = {
            "total": total_h,
            "fail": fail_h,
            "ok": ok_h,
            "success_rate": success_rate,
            "fail_rate": fail_h / total_h if total_h > 0 else np.nan,
        }

        hub_df = df[df["集配站"] == hub].copy()
        hub_fail_df = fail_df[fail_df["集配站"] == hub].copy()
        # === 新增：by 配送站的问题明细（保留配送站为空） ===
        station_total = (
            hub_df["配送站"]
            .value_counts()
            .rename_axis("配送站")
            .reset_index(name="配送站总单量")
        )
        
        station_summary = (
            hub_fail_df
            .groupby(
                ["配送站", "链路问题归因", "主要责任方"],
                dropna=False
            )
            .size()
            .reset_index(name="问题单量")
        )
        
        station_summary = station_summary.merge(station_total, on="配送站", how="left")
        station_summary["占比_numeric"] = station_summary["问题单量"] / station_summary["配送站总单量"].replace(0, np.nan)
        station_summary = station_summary.sort_values(
            ["配送站", "占比_numeric"],
            ascending=[True, False]
        )
        station_summary["占配送站总单量比"] = (
            station_summary["占比_numeric"] * 100
        ).round(2).fillna(0).astype(str) + "%"
        station_summary = station_summary.drop(columns=["占比_numeric"])
        
        station_summary_display = station_summary.copy()

        # 让相邻重复的“配送站”显示为空（只保留第一行）
        same_as_prev = station_summary_display["配送站"].eq(station_summary_display["配送站"].shift())
        
        # 处理 NaN：如果当前和上一行都是空，也算重复
        same_nan_as_prev = station_summary_display["配送站"].isna() & station_summary_display["配送站"].shift().isna()
        
        station_summary_display.loc[same_as_prev | same_nan_as_prev, "配送站"] = ""
        hub_sta_summary[hub] = station_summary_display
    
    # ===== 4. Output results as an excel =====
    output_file = f"SLA_分析完成.xlsx"
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # 明细表
        fail_df.to_excel(writer, sheet_name="明细", index=False)
        
        ws = writer.sheets["明细"]
        ws.set_column(0, fail_df.shape[1]-1, 18)
        ws.set_column("AJ:AJ", 35)
        
        # 整体问题归因统计表
        overall_info.to_excel(writer, sheet_name="整体问题归因统计", index=False, startrow=0)
        start_row = len(overall_info) + 2
        summary_all.to_excel(writer, sheet_name="整体问题归因统计", index=False, startrow=start_row)
        start_row = start_row + len(summary_all) + 2
        hub_overall.to_excel(writer, sheet_name="整体问题归因统计", index=False, startrow=start_row)
    
        ws = writer.sheets["整体问题归因统计"]
        ws.set_column(0, max(summary_all.shape[1], overall_info.shape[1]) - 1, 18)
        ws.set_column("A:A", 35)
    
        # By客户问题归因表
        for client in sorted(df["客户"].unique()):
            sub = summary_by_client[summary_by_client["客户"] == client].copy()
    
            # No Fail Order
            if sub.empty:
                sla_info = client_sla_summary.get(client, None)
                info_rows = [
                    ["客户", client],
                    ["总单量", sla_info["total"] if sla_info else 0],
                    ["未达标单量", 0],
                    ["未达标率", "0.00%"],
                    ["目标达成率", f"{sla_info['target']*100:.0f}%" if sla_info and not np.isnan(sla_info["target"]) else "未配置"],
                    ["是否达标", "达标 ✔" if sla_info and sla_info["meet_target"] else "未达标 ❌"]
                ]
                info_df = pd.DataFrame(info_rows, columns=["指标", "值"])
                info_df.to_excel(writer, sheet_name=client, index=False)
                ws = writer.sheets[client]
                ws.set_column(0, info_df.shape[1]-1, 18)
                continue
    
            # Have Fail Order
            sla_info = client_sla_summary[client]
    
            info_rows = [
                ["客户", client],
                ["总单量", sla_info["total"]],
                ["未达标单量", sla_info["fail"]],
                ["未达标率", f"{sla_info['fail_rate']*100:.2f}%"],
                ["目标达成率", f"{sla_info['target']*100:.0f}%"],
                ["是否达标", "达标 ✔" if sla_info["meet_target"] else "未达标 ❌"]
            ]
            info_df = pd.DataFrame(info_rows, columns=["指标", "值"])
            sub = sub.drop(columns=["客户"])
    
            sheet_name = client
            info_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
            start_row = len(info_df) + 2   # 空一行
            sub.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
    
            ws = writer.sheets[sheet_name]
            ws.set_column(0, max(info_df.shape[1], sub.shape[1]) - 1, 18)
            ws.set_column("A:A", 35)
    
        # By hub问题归因表
        hubs = (
            df["集配站"]
            .dropna()
            .astype(str)
            .unique()
        )
        for hub in sorted(hubs):
            sub = summary_by_hub[summary_by_hub["集配站"] == hub].copy()
    
            # No Fail Order
            if sub.empty:
                sla_info = hub_sla_summary.get(hub, None)
                info_rows = [
                    ["集配站", hub],
                    ["总单量", sla_info["total"] if sla_info else 0],
                    ["未达标单量", 0],
                    ["未达标率", "0.00%"]
                ]
                info_df = pd.DataFrame(info_rows, columns=["指标", "值"])
                info_df.to_excel(writer, sheet_name=hub, index=False)
                ws = writer.sheets[hub]
                ws.set_column(0, info_df.shape[1] - 1, 18)
                continue
    
            # Have Fail Order
            sla_info = hub_sla_summary[hub]
    
            info_rows = [
                ["集配站", hub],
                ["总单量", sla_info["total"]],
                ["未达标单量", sla_info["fail"]],
                ["未达标率", f"{sla_info['fail_rate']*100:.2f}%"],
            ]
            info_df = pd.DataFrame(info_rows, columns=["指标", "值"])
            sub = sub.drop(columns=["集配站"])
    
            sheet_name = hub
            info_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
            start_row = len(info_df) + 2   # 空一行
            sub.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
            start_row = start_row + len(sub) + 2
            hub_sta_summary[hub].to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
    
            ws = writer.sheets[sheet_name]
            ws.set_column(0, max(info_df.shape[1], sub.shape[1]), 18)
            ws.set_column("A:A", 35)
    
    return {
        "filename": output_file,
        "output_bytes": output.getvalue()
    }