from totalasia_matcher import run_stock_check, MatchConfig

cfg = MatchConfig(
    min_common_words_primary=3,
    min_common_words_relaxed=2,
    product_pad_width=6   # 统一补齐到6位
)

stock_check, matched_map, unmatched = run_stock_check(
    supplier_path="hh001.xlsx",
    erp_path="ERP hh001.xlsx",
    output_dir="./output",
    cfg=cfg,
    export_unmatched=True,
    fill_missing_qty_zero=True  # 启用：Qty 空缺填 0
)

print(stock_check.shape)
print("Unmatched:", len(unmatched))

#import pandas as pd
#df = pd.read_excel("./output/Stock Check.xlsx", dtype={"Product": str})
#print(df["Product"].head(10).tolist())