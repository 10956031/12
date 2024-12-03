#!/usr/bin/env python
# coding: utf-8

# In[4]:


import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from tkinter import Tk, Label, Button, filedialog, Frame
from tkinter.ttk import Treeview
from matplotlib.cm import tab10

# 設置中文字體為微軟正黑體，中文顯示
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']  # 微軟正黑體
plt.rcParams['axes.unicode_minus'] = False  # 支援負號顯示


class InventoryManagementApp:
    def __init__(self, master):
        self.master = master
        self.master.title("進銷存系統")
        self.master.geometry("1800x1080")

        # 狀態標籤
        self.status_label = Label(master, text="", fg="blue")
        self.status_label.grid(row=0, column=0, columnspan=3, pady=5)

        # 分區 Frame
        self.left_frame = Frame(master)
        self.left_frame.grid(row=1, column=0, padx=20, pady=20, sticky="n")

        self.center_frame = Frame(master)
        self.center_frame.grid(row=1, column=1, padx=20, pady=20, sticky="n")

        self.right_frame = Frame(master)
        self.right_frame.grid(row=1, column=2, padx=20, pady=20, sticky="n")

        # 左側按鈕
        Button(self.left_frame, text="下載帶範例的資料庫", command=self.download_template).pack(pady=10)
        Button(self.left_frame, text="上傳資料庫檔案", command=self.upload_file).pack(pady=10)
        Button(self.left_frame, text="生成銷售趨勢圖", command=self.generate_sales_trend).pack(pady=10)
        
        # 中央按鈕
        Button(self.center_frame, text="生成銷售金額堆疊圖", command=self.generate_sales_stack).pack(pady=10)
        Button(self.center_frame, text="生成每周進貨報表", command=self.show_weekly_purcurement_report).pack(pady=10)
        Button(self.center_frame, text="生成庫存趨勢圖", command=self.generate_inventory_trend).pack(pady=10)  # 新增的按鈕

        # 右側按鈕
        Button(self.right_frame, text="生成庫存報表", command=self.show_inventory_report).pack(pady=10)  
        Button(self.right_frame, text="生成利潤表", command=self.show_profit_table).pack(pady=10)
        Button(self.right_frame, text="供應商查詢", command=self.show_supplier_report).pack(pady=10)

        # 表格/圖表區域
        self.chart_frame = Frame(master)
        self.chart_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="nsew")

        # 初始化檔案路徑與數據
        self.filepath = None
        self.orders_df = None
        self.purcurement_df = None
        self.inventory_df = None
        self.products_df = None
        self.suppliers_df = None

    def update_status(self, message, color="blue"):
        """更新狀態標籤"""
        self.status_label.config(text=message, fg=color)

    def upload_file(self):
        """上傳資料庫"""
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.filepath:
            try:
                self.orders_df = pd.read_excel(self.filepath, sheet_name=0)
                self.purcurement_df = pd.read_excel(self.filepath, sheet_name=1)
                self.inventory_df = pd.read_excel(self.filepath, sheet_name=2)
                self.products_df = pd.read_excel(self.filepath, sheet_name=3)
                self.suppliers_df = pd.read_excel(self.filepath, sheet_name=4)
                self.update_status("檔案已成功上傳！", "green")
            except Exception as e:
                self.update_status(f"檔案格式錯誤：{e}", "red")

    def download_template(self):
        """下載資料庫模板"""
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
        initialfile="進銷存資料庫.xlsx")  # 設置預設檔名  
        if save_path:
            try:
                # 模板資料
                data = {
                    "orders": pd.DataFrame({
                        "OrderID": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
                        "ProductID": ["L", "G", "S", "L", "G", "S", "L", "G", "S", "S", "G", "S", "L", "G", "S"],
                        "Quantity": [4, 2, 3, 3, 5, 3, 4,3, 3, 2, 3, 4, 3, 4, 3],
                        "Customer": ["M", "F", "F", "M", "F", "M", "M", "M", "F", "F", "M", "F", "M", "F", "M"],
                        "Week": [0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4]
                    }),
                    "purcurement": pd.DataFrame({
                        "PurcurementID": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
                        "ProductID": ["L", "G", "S", "L", "G", "S", "L", "G", "S", "L", "G", "S", "L", "G", "S"],
                        "Quantity": [5, 4, 3, 3, 3, 4, 2, 3, 2, 3, 4, 4, 3, 3, 4],
                        "Supplier": ["供應商A", "供應商B", "供應商C", "供應商A", "供應商B", "供應商C", 
                        "供應商A", "供應商B", "供應商C", "供應商A", "供應商B", "供應商C", 
                        "供應商A", "供應商B", "供應商C"],
                        "Week": [0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4]
                    }),
                    "inventory": pd.DataFrame({
                        "ProductID": ["L", "G", "S"],
                        "Quantity": [6, 6, 7],
                        "Week": [0, 0, 0]
                    }),
                    "products": pd.DataFrame({
                        "ProductID": ["L", "G", "S"],
                        "ProductName": ["龍蝦", "鮭魚", "蝦米"],
                        "CategoryID": ["A", "B", "C"],
                        "Unit": ["尾", "公克", "公克"],
                        "Price": [4, 3, 1.5],
                        "Cost": [2, 2, 1]
                    }),
                    "suppliers": pd.DataFrame({
                        "ProductID": ["L", "G", "S"],
                        "SupplierName": ["南寮水產", "基隆水產之家", "馬力水產"],
                        "ContactName": ["李大成", "王天賜", "吳文新"],
                        "Phone": ["0922732525", "0908223534", "0937200311"],
                        "Email": ["supplierA@example.com", "supplierB@example.com", "supplierC@example.com"],
                        "Address": ["地址A", "地址B", "地址C"]
                    })
                }

                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    for sheet_name, df in data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

                self.update_status("範例已成功生成！", "green")
            except Exception as e:
                self.update_status(f"範例生成失敗：{e}", "red")

    def show_table_in_ui(self, data, columns):
        """在介面中顯示表格，並增加保存按鈕"""
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

        if data.empty:
            self.update_status("無數據顯示！", "red")
            return

        # 建立 Treeview 表格
        tree = Treeview(self.chart_frame, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")
        tree.pack(fill="both", expand=True)

        for _, row in data.iterrows():
            tree.insert("", "end", values=tuple(row))

        # 增加保存按鈕
        save_button = Button(self.chart_frame, text="另存為 Excel 檔案", command=lambda: self.save_table_as_excel(data))
        save_button.pack(side="bottom", pady=5)

    def show_chart_in_ui(self, figure):
        """將 Matplotlib 圖表顯示於 Tkinter 介面中，並增加保存按鈕"""
        for widget in self.chart_frame.winfo_children():
            widget.destroy()

        # 調整圖表大小，防止被截斷
        figure.set_size_inches(8, 6)  # 調整圖表寬度和高度

        canvas = FigureCanvasTkAgg(figure, master=self.chart_frame)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(fill="both", expand=True)
        canvas.draw()

        # 增加保存按鈕
        save_button = Button(self.chart_frame, text="另存為 PNG 檔案", command=lambda: self.save_chart_as_png(figure))
        save_button.pack(side="bottom", pady=5)

    def save_chart_as_png(self, figure):
        """將圖表另存為 PNG 檔案"""
        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if file_path:
            try:
                figure.savefig(file_path)
                self.update_status("圖片已成功保存！", "green")
            except Exception as e:
                self.update_status(f"圖片保存失敗：{e}", "red")
                
    def save_table_as_excel(self, data):
        """將表格數據另存為 Excel 檔案"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                data.to_excel(file_path, index=False)
                self.update_status("Excel 檔案已成功保存！", "green")
            except Exception as e:
                self.update_status(f"Excel 檔案保存失敗：{e}", "red")                

    def generate_sales_trend(self):
        """產生銷售趨勢圖"""
        if self.orders_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        merged_data = self.orders_df.merge(self.products_df, on="ProductID")
        trend_data = merged_data.groupby(["Week", "ProductName"]).sum().reset_index()

        fig, ax = plt.subplots()
        for product in trend_data["ProductName"].unique():
            product_data = trend_data[trend_data["ProductName"] == product]
            ax.plot(product_data["Week"], product_data["Quantity"], label=product)

        ax.set_title("銷售趨勢圖")
        ax.set_xlabel("週")
        ax.set_ylabel("銷售數量")
        ax.legend()
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))
        # 設置 Y 軸的起始值為 0
        ax.set_ylim(bottom=0)        

        self.show_chart_in_ui(fig)

    def generate_sales_stack(self):
        """產生銷售金額堆疊圖，並提供另存為 PNG 功能"""
        if self.orders_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 合併銷售和產品資料
        merged_data = self.orders_df.merge(self.products_df, on="ProductID")
        # 計算每週每個產品的銷售金額
        stack_data = merged_data.groupby(["Week", "ProductName"]).apply(
            lambda x: (x["Quantity"] * x["Price"]).sum()
        ).unstack()
        
        # 繪製堆疊圖
        fig, ax = plt.subplots()
        stack_data.plot(kind="bar", stacked=True, ax=ax, colormap=tab10)
        ax.set_title("銷售金額堆疊圖")
        ax.set_xlabel("週")
        ax.set_ylabel("銷售金額")
        ax.legend(title="產品名稱")

        # 在 UI 中顯示圖表並附加保存按鈕
        self.show_chart_in_ui(fig)

    def show_weekly_purcurement_report(self):
        """產生每周各產品的進貨報表，並提供另存為 Excel 功能"""
        if self.purcurement_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 合併資料並生成報表
        merged_data = self.purcurement_df.merge(self.products_df, on="ProductID")
        purcurement_report = merged_data.groupby(["Week", "ProductName"]).sum()["Quantity"].reset_index()

        # 在介面中顯示表格並附加保存按鈕
        self.show_table_in_ui(purcurement_report, ["Week", "ProductName", "Quantity"])

    def show_profit_table(self):
        """產生利潤表，並提供另存為 Excel 的功能"""
        if self.orders_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 合併銷售和產品數據，計算利潤
        merged_data = self.orders_df.merge(self.products_df, on="ProductID")
        merged_data["Profit"] = (merged_data["Price"] - merged_data["Cost"]) * merged_data["Quantity"]
        profit_table = merged_data.groupby("ProductName")[["Quantity", "Profit"]].sum().reset_index()

        # 在介面中顯示利潤表並附加保存按鈕
        self.show_table_in_ui(profit_table, ["ProductName", "Quantity", "Profit"])

    def show_inventory_report(self):
        """產生庫存報表，並新增產品名稱(ProductName)欄"""
        if self.inventory_df is None or self.orders_df is None or self.purcurement_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 初始化庫存數據
        inventory_report = []
        for product_id in self.inventory_df["ProductID"].unique():
            product_inventory = self.inventory_df[self.inventory_df["ProductID"] == product_id]
            product_orders = self.orders_df[self.orders_df["ProductID"] == product_id]
            product_purcurement = self.purcurement_df[self.purcurement_df["ProductID"] == product_id]
            previous_inventory = product_inventory.iloc[0]["Quantity"]

            for week in range(1, max(product_inventory["Week"].max(), product_orders["Week"].max(), product_purcurement["Week"].max()) + 1):
                weekly_sales = product_orders[product_orders["Week"] == week]["Quantity"].sum()
                weekly_purcurement = product_purcurement[product_purcurement["Week"] == week]["Quantity"].sum()
                current_inventory = previous_inventory + (weekly_purcurement - weekly_sales)

                inventory_report.append({
                    "ProductID": product_id,
                    "Week": week,
                    "PreviousInventory": previous_inventory,
                    "Sales": weekly_sales,
                    "Purcurement": weekly_purcurement,
                    "CurrentInventory": current_inventory
                })
                previous_inventory = current_inventory

        # 將庫存數據轉為 DataFrame
        inventory_report_df = pd.DataFrame(inventory_report)

        # 確保 ProductID 的數據類型一致，並與 products_df 合併
        inventory_report_df["ProductID"] = inventory_report_df["ProductID"].astype(str).str.strip()
        self.products_df["ProductID"] = self.products_df["ProductID"].astype(str).str.strip()

        # 合併以新增產品名稱
        inventory_report_with_name = inventory_report_df.merge(self.products_df[["ProductID", "ProductName"]], on="ProductID", how="left")

        # 在介面中顯示報表，新增 ProductName 欄位
        columns_to_display = ["ProductID", "ProductName", "Week", "PreviousInventory", "Sales", "Purcurement", "CurrentInventory"]
        self.show_table_in_ui(inventory_report_with_name[columns_to_display], columns_to_display)

    def generate_inventory_trend(self):
        """產生每周庫存總金額趨勢圖"""
        if self.inventory_df is None or self.orders_df is None or self.purcurement_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 初始化庫存數據
        inventory_report = []
        for product_id in self.inventory_df["ProductID"].unique():
            product_inventory = self.inventory_df[self.inventory_df["ProductID"] == product_id]
            product_orders = self.orders_df[self.orders_df["ProductID"] == product_id]
            product_purcurement = self.purcurement_df[self.purcurement_df["ProductID"] == product_id]
            previous_inventory = product_inventory.iloc[0]["Quantity"]

            # 獲取產品成本
            product_cost = self.products_df.loc[self.products_df["ProductID"] == product_id, "Cost"].values[0]

            for week in range(1, max(product_inventory["Week"].max(), product_orders["Week"].max(), product_purcurement["Week"].max()) + 1):
                weekly_sales = product_orders[product_orders["Week"] == week]["Quantity"].sum()
                weekly_purcurement = product_purcurement[product_purcurement["Week"] == week]["Quantity"].sum()
                current_inventory = previous_inventory + (weekly_purcurement - weekly_sales)

                inventory_report.append({
                    "ProductID": product_id,
                    "Week": week,
                    "CurrentInventory": current_inventory,
                    "InventoryValue": current_inventory * product_cost
                })
                previous_inventory = current_inventory

        # 將庫存數據轉為 DataFrame
        inventory_report_df = pd.DataFrame(inventory_report)

        # 計算每周的庫存總金額
        total_inventory_value = inventory_report_df.groupby("Week")["InventoryValue"].sum().reset_index()

        # 繪製庫存總金額趨勢圖
        fig, ax = plt.subplots()

        ax.plot(
            total_inventory_value["Week"],
            total_inventory_value["InventoryValue"],
            label="庫存總金額",
            marker="o"
        )

        # 設置圖表標題和標籤
        ax.set_title("每周庫存總金額趨勢圖")
        ax.set_xlabel("週")
        ax.set_ylabel("庫存總金額")

        # 設置 X 軸為整數
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))

        # 設置 Y 軸範圍
        if not total_inventory_value["InventoryValue"].empty:
            ax.set_ylim(0, total_inventory_value["InventoryValue"].max() * 1.1)

        # 添加圖例
        ax.legend()

        # 在介面中顯示圖表
        self.show_chart_in_ui(fig)
        
    def show_supplier_report(self):
        """產生供應商報表，並新增產品名稱(ProductName)欄"""
        if self.suppliers_df is None or self.products_df is None:
            self.update_status("請先上傳資料庫！", "red")
            return

        # 確保 ProductID 的數據類型一致（轉為字符串類型）
        self.suppliers_df["ProductID"] = self.suppliers_df["ProductID"].astype(str).str.strip()
        self.products_df["ProductID"] = self.products_df["ProductID"].astype(str).str.strip()

        # 調試：檢查 ProductID 的唯一值是否一致
        missing_ids = set(self.suppliers_df["ProductID"]) - set(self.products_df["ProductID"])
        if missing_ids:
            print(f"以下 ProductID 無匹配：{missing_ids}")

        # 合併供應商與產品數據
        supplier_report = self.suppliers_df.merge(self.products_df, on="ProductID", how="left")

        # 選擇需要顯示的欄位
        columns_to_display = ["ProductID", "ProductName", "SupplierName", "ContactName", "Phone", "Email", "Address"]

        # 在介面中顯示報表
        self.show_table_in_ui(supplier_report[columns_to_display], columns_to_display)


# 主程式執行入口
if __name__ == "__main__":
    root = Tk()
    app = InventoryManagementApp(root)
    root.mainloop()


# In[ ]:




