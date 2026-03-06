import pandas as pd


class CompanyConfigManager:
    def __init__(self, excel_path="config/company_steps.xlsx"):
        self.df = pd.read_excel(excel_path, dtype=str).fillna('')
        # 确保按 Excel 顺序排列
        self.df = self.df.reset_index(drop=True)

    def get_steps(self, company_name: str):
        """返回该公司的步骤列表（按顺序）"""
        return self.df[self.df['Company'] == company_name]['Step'].tolist()

    def get_step_config(self, company_name: str, step_name: str):
        """返回该步骤的配置：field_name + hint"""
        row = self.df[(self.df['Company'] == company_name) & (self.df['Step'] == step_name)]
        if row.empty:
            return {"hint": ""}
        return {
            "hint": row.iloc[0]['Hint']
        }
