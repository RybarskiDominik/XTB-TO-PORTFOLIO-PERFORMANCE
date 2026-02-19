import pandas as pd
import datetime
import logging
import re
import os

logger = logging.getLogger(__name__)

class CashOperationXLSXReader:
    def __init__(self, xlsx_path: str, sheet_index: int = 3):
        self.xlsx_path = xlsx_path
        self.sheet_index = sheet_index
        self.account_currency = None
        self.df = None

        self.header = {}

        self.all_operations = pd.DataFrame()
        self.operations = pd.DataFrame()
        self.cfd_operations = pd.DataFrame()
        self.operations_on_account = pd.DataFrame()
        self.operations_on_stocks = pd.DataFrame()

    # ---------- LOAD XLSX ----------
    def load_sheet(self):
        self.df = pd.read_excel(
            self.xlsx_path,
            sheet_name=self.sheet_index,
            header=None
        )

    # ---------- READ HEADER ----------
    def read_header(self) -> dict:
        if self.df is None:
            self.load_sheet()

        header = {}

        # --- Name / Account / Currency ---
        r = self._find_row_with("Name and surname")
        if r is not None:
            row = self.df.iloc[r]
            idx = row[row == "Name and surname"].index[0]

            values = (
                self.df.iloc[r + 1, idx:]
                .dropna()
                .tolist()
            )

            header["Name and surname"] = values[0]
            header["Account"] = values[1] if len(values) > 1 else None
            header["Currency"] = values[2] if len(values) > 2 else None

            self.account_currency = header["Currency"]

        # --- Balance block ---
        r = self._find_row_with("Balance")
        if r is not None:
            row = self.df.iloc[r]
            idx = row[row == "Balance"].index[0]

            values = (
                self.df.iloc[r + 1, idx:]
                .dropna()
                .tolist()
            )

            keys = [
                "Balance",
                "Equity",
                "Margin",
                "Free margin",
                "Margin level"
            ]

            for k, v in zip(keys, values):
                header[k] = self._num(v)

        self.header = header
        return header

    def _find_row_with(self, text: str):
        for i in range(len(self.df)):
            if text in self.df.iloc[i].astype(str).values:
                return i
        return None

    # ---------- READ TOTAL ----------
    def read_total(self) -> dict:
        if self.df is None:
            self.load_sheet()

        for i in range(len(self.df)):
            row = self.df.iloc[i]
            if str(row[1]).strip() == "Total":
                print(row)
                return {
                    "Total": self._num(row[6]),
                    "Currency": row[7]
                }

        return {"Total": None, "Currency": None}

    # ---------- READ TABLE OPERATIONS ----------
    def read_table(self, columns: list) -> pd.DataFrame:
        if self.df is None:
            self.load_sheet()

        header_row = None
        col_index = {}

        # Search for header row
        for i in range(len(self.df)):
            row = self.df.iloc[i].astype(str).str.strip().tolist()
            matches = [c for c in columns if c in row]

            if len(matches) >= max(3, len(columns) // 2):
                header_row = i
                for c in columns:
                    if c in row:
                        col_index[c] = row.index(c)
                break

        if header_row is None:
            raise ValueError("Header row not found.")

        data = []

        for i in range(header_row + 1, len(self.df)):
            row = self.df.iloc[i]

            if row.astype(str).str.contains("TOTAL", case=False).any():
                break

            record = {}
            empty = True

            for c, idx in col_index.items():
                val = row.iloc[idx]
                if not pd.isna(val):
                    empty = False
                record[c] = val

            if not empty:
                data.append(record)

        self.operations = pd.DataFrame(data)

        return self.operations

    # ---------- OPERATIONS HISTORY NORMALIZATION ----------
    def normalize_operations_history(self, amount=False, lang="EN"):
        """Rename columns and map operation types."""

        # --- TYPE MAP ---
        type_map = {
            "deposit": "Deposit",
            "Stock purchase": "Buy",
            "close trade": "close trade",
            "Stock sale": "Sell",
            "DIVIDENT": "Dividend",
            "withdrawal": "Withdrawal",
            "Withholding Tax": "Taxes",
            "Free-funds Interest": "Interest",
            "Free-funds Interest Tax": "Taxes",
            "transfer": "transfer",
            # "transfer": "Transfer (Inbound)",
            # "transfer": "Transfer (Outbound)"
        }

        # --- COLUMN MAP ---
        column_map = {
            "ID": "ID",
            "Type": "Type",
            "Time": "Date",
            "Comment": "Note",
            "Symbol": "Ticker Symbol",
            "Amount": "Amount",
            "Position": "Position",
            "Volume": "Shares",
            "Open time": "Date",
            "Open price": "Value",
            "Close time": "Date",
            "Close price": "Value",
            "Purchase": "Purchase",
            "value": "Value",
            "SL": "SL",
            "TP": "TP",
            "Margin": "Margin",
            "Commission": "Commission",
            "Swap": "Swap",
            "Rollover": "Rollover",
            "Gross P/L": "Gross P/L",
        }

        # Rename columns
        self.operations.rename(columns=column_map, inplace=True)

        # Map Type values
        if "Type" in self.operations.columns:
            self.operations['Type'] = self.operations['Type'].map(type_map).fillna(self.operations['Type'])

        # --- TRANSFER: NEGATIVE -> WITHDRAWAL ---
        if "Type" in self.operations.columns and "Amount" in self.operations.columns:
            mask_transfer = self.operations["Type"] == "transfer"

            self.operations.loc[mask_transfer & (self.operations["Amount"] > 0), "Type"] = "Deposit"
            self.operations.loc[mask_transfer & (self.operations["Amount"] < 0), "Type"] = "Withdrawal"

            # self.operations.loc[mask_transfer & (self.operations["Amount"] > 0), "Note"] = "Transfer Inbound"
            # self.operations.loc[mask_transfer & (self.operations["Amount"] < 0), "Note"] = "Transfer Outbound"

        # --- HANDLE CFD CLOSE TRADE ---
        if "Type" in self.operations.columns and "Ticker Symbol" in self.operations.columns:
            def handle_cfd(row):
                if row["Type"] == "close trade":
                    ticker = str(row["Ticker Symbol"]).strip()
                    # Check: plain code, no dots or other characters
                    if re.match(r"^[A-Z0-9]+$", ticker):
                        # CFD Profit or Loss
                        if row["Amount"] >= 0:
                            row["Note"] = "Profit CFD"
                            row["Type"] = "Deposit"
                        else:
                            row["Note"] = "Loss CFD"
                            row["Type"] = "Withdrawal"
                return row

            self.operations = self.operations.apply(handle_cfd, axis=1)

        # --- SKIP CLOSE TRADE ---
        self.operations = self.operations[self.operations["Type"] != "close trade"]

        # --- CURRENCY FIX ---
        self.operations["Transaction Currency"] = self.account_currency
        self.operations["Currency Gross Amount"] = self.account_currency
        self.operations["Cash Account"] = self.account_currency
        self.operations["Securities Account"] = f"XTB {self.account_currency}"

        # --- SHARES + PRICE ---
        self.operations = self.add_quantity_and_price(self.operations)

        # Fill text columns
        self.operations["Ticker Symbol"] = self.operations["Ticker Symbol"].fillna("")
        self.operations["Note"] = self.operations["Note"].fillna("")

        # --- Value ← Amount (only when Value is empty) ---
        self.operations["Value"] = self.operations["Value"].replace("", pd.NA)
        mask = self.operations["Value"].isna()
        self.operations.loc[mask, "Value"] = self.operations.loc[mask, "Amount"]

        # Format date and amount
        self.operations["Date"] = pd.to_datetime(
            self.operations["Date"]
        ).dt.strftime("%Y-%m-%dT%H:%M")

        self.operations["Amount"] = self.operations["Amount"].apply(
            lambda x: f"{x:.4f}" if pd.notna(x) else ""
        )

        return self.operations

    # ---------- ADD QUANTITY AND PRICE ----------
    def add_quantity_and_price(self, df: pd.DataFrame) -> pd.DataFrame:
        df["Shares"] = ""
        df["Gross Amount"] = ""

        for idx, row in df.iterrows():
            note = str(row.get("Note", ""))

            if "OPEN BUY" in note or "CLOSE BUY" in note:
                if "/" in note:
                    index1 = note.index("/")
                    index2 = note.index("@")
                else:
                    index1 = index2 = note.index("@")

                # Shares = substring(9, index1 - 9)
                shares = note[9:index1].strip()

                # Gross Amount = substring from index2 + 2
                gross_amount = note[index2 + 2:].strip()

                df.at[idx, "Shares"] = shares
                df.at[idx, "Gross Amount"] = gross_amount

                # Value = Shares * Gross Amount
                try:
                    df.at[idx, "Value"] = float(shares) * float(gross_amount)
                    # df.at[idx, "Value"] = abs(row.get("Amount", 0))
                except ValueError:
                    pass

        self.operations = df

        return self.operations

    # ---------- OPEN OPERATIONS NORMALIZATION ----------
    def normalize_open_operations(self, amount=False, lang="EN"):
        """Rename columns and map operation types."""

        # --- TYPE MAP ---
        type_map = {
            "deposit": "Deposit",
            "Stock purchase": "Buy",
            "close trade": "close trade",
            "Stock sale": "Sell",
            "DIVIDENT": "Dividend",
            "withdrawal": "Withdrawal",
            "Withholding Tax": "Taxes",
            "Free-funds Interest": "Interest",
            "Free-funds Interest Tax": "Taxes",
            "transfer": "transfer",
            # "transfer": "Transfer (Inbound)",
            # "transfer": "Transfer (Outbound)"
        }

        # --- COLUMN MAP ---
        column_map = {
            "ID": "ID",
            "Type": "Type",
            "Time": "Date",
            "Comment": "Note",
            "Symbol": "Ticker Symbol",
            "Amount": "Amount",
            "Position": "Position",
            "Volume": "Shares",
            "Open time": "Date",
            "Open price": "Value",
            "Close time": "Date",
            "Purchase value": "Gross Amount",
            # "Close price": "Value",
            "Purchase": "Purchase",
            # "value": "Value",
            "SL": "SL",
            "TP": "TP",
            "Margin": "Margin",
            "Commission": "Commission",
            "Swap": "Swap",
            "Rollover": "Rollover",
            "Gross P/L": "Gross P/L",
        }

        # Rename columns
        self.operations.rename(columns=column_map, inplace=True)

        # Map Type values
        if "Type" in self.operations.columns:
            self.operations['Type'] = self.operations['Type'].map(type_map).fillna(self.operations['Type'])

        # --- CURRENCY FIX ---
        self.operations["Transaction Currency"] = self.account_currency
        self.operations["Currency Gross Amount"] = self.account_currency
        self.operations["Cash Account"] = self.account_currency
        self.operations["Securities Account"] = f"XTB {self.account_currency}"

        # Fill text columns
        self.operations["Ticker Symbol"] = self.operations["Ticker Symbol"].fillna("")
        self.operations["Note"] = self.operations["Note"].fillna("")

        self.operations["Value"] = self.operations["Shares"] * self.operations["Value"]

        # Format date
        self.operations["Date"] = pd.to_datetime(
            self.operations["Date"]
        ).dt.strftime("%Y-%m-%dT%H:%M")

        return self.operations

    # ---------- CLOSED OPERATIONS NORMALIZATION ----------
    def normalize_closed_operations(self, amount=False, lang="EN"):
        """Rename columns and map operation types."""

        # --- TYPE MAP ---
        type_map = {
            "deposit": "Deposit",
            "Stock purchase": "Buy",
            "close trade": "close trade",
            "Stock sale": "Sell",
            "DIVIDENT": "Dividend",
            "withdrawal": "Withdrawal",
            "Withholding Tax": "Taxes",
            "Free-funds Interest": "Interest",
            "Free-funds Interest Tax": "Taxes",
            "transfer": "transfer",
            # "transfer": "Transfer (Inbound)",
            # "transfer": "Transfer (Outbound)"
        }

        # --- COLUMN MAP ---
        column_map = {
            "ID": "ID",
            "Type": "Type",
            "Time": "Date",
            "Comment": "Note",
            "Symbol": "Ticker Symbol",
            "Amount": "Amount",
            "Position": "Position",
            "Volume": "Shares",
            # "Open time": "Date",
            # "Open price": "Value",
            # "Close time": "Date",
            # "Close price": "Value",
            "Purchase": "Purchase",
            "value": "Value",
            "SL": "SL",
            "TP": "TP",
            "Margin": "Margin",
            "Commission": "Commission",
            "Swap": "Swap",
            "Rollover": "Rollover",
            "Gross P/L": "Gross P/L",
        }

        # Rename columns
        self.operations.rename(columns=column_map, inplace=True)

        # Map Type values
        if "Type" in self.operations.columns:
            self.operations['Type'] = self.operations['Type'].map(type_map).fillna(self.operations['Type'])

        self.split_cfd_and_stocks()

        # --- HANDLE CFD CLOSE TRADE ---
        if "Type" in self.operations.columns and "Ticker Symbol" in self.operations.columns:
            def handle_cfd(row):
                ticker = str(row["Ticker Symbol"]).strip()
                # Check: plain code, no dots or other characters
                if re.match(r"^[A-Z0-9]+$", ticker):
                    # CFD Profit or Loss
                    if row["Gross P/L"] >= 0:
                        row["Note"] = "Profit CFD on: " + ticker + " on " + str(row["Close time"])
                        row["Type"] = "Deposit"
                        row["Value"] = row["Gross P/L"]
                    else:
                        row["Note"] = "Loss CFD on: " + ticker + " on " + str(row["Close time"])
                        row["Type"] = "Withdrawal"
                        row["Value"] = row["Gross P/L"]

                    #print(f"Updated Type: {row['Type']}, Note: {row['Note']}, Value: {row['Value']}")

                return row

            self.cash_flow_cfd_operations = self.cfd_operations.apply(handle_cfd, axis=1)

            if not self.cash_flow_cfd_operations.empty:
                self.cash_flow_cfd_operations = self.cash_flow_cfd_operations[["Position", "Type", "Close time", "Value", "Note"]].copy()

        self.open_operations = self.operations[["Position", "Ticker Symbol", "Type", "Shares", "Open time", "Open price", "Purchase value", "Note"]].copy()
        self.closed_operations = self.operations[["Position", "Ticker Symbol", "Type", "Shares", "Close time", "Close price", "Sale value", "Note"]].copy()
        self.closed_operations["Type"] = "Sell"

        self.open_operations.rename(columns={
            "Open time": "Date",
            "Open price": "Value",
            "Purchase value": "Gross Amount"
        }, inplace=True)

        self.closed_operations.rename(columns={
            "Close time": "Date",
            "Close price": "Value",
            "Sale value": "Gross Amount"
        }, inplace=True)

        self.cash_flow_cfd_operations.rename(columns={
            "Close time": "Date",
        }, inplace=True)

        self.operations = pd.concat([self.open_operations, self.closed_operations], ignore_index=True)

        self.operations["Value"] = self.operations["Shares"] * self.operations["Value"]

        self.operations = pd.concat([self.operations, self.cash_flow_cfd_operations], ignore_index=True)

        #print(self.open_operations)
        #print(self.closed_operations)
        #print(self.cash_flow_cfd_operations)

        #print(self.operations)

        # --- CURRENCY FIX ---
        self.operations["Transaction Currency"] = self.account_currency
        self.operations["Currency Gross Amount"] = self.account_currency
        self.operations["Cash Account"] = self.account_currency
        self.operations["Securities Account"] = f"XTB {self.account_currency}"

        # Fill text columns
        self.operations["Ticker Symbol"] = self.operations["Ticker Symbol"].fillna("")
        self.operations["Note"] = self.operations["Note"].fillna("")

        self.operations["Date"] = pd.to_datetime(self.operations["Date"]).dt.strftime("%Y-%m-%dT%H:%M")

        return self.operations

    # ---------- HELPERS ----------
    @staticmethod
    def _num(val):  # Convert a numeric string to float, handling comma as decimal separator.
        if pd.isna(val):
            return None
        return float(str(val).replace(",", "."))

    # ---------- STRIP EXCHANGE SUFFIX ----------
    def strip_ticker_suffix(self):
        """
        Removes everything after the dot in the 'Ticker Symbol' column.
        Example: MSFT.US -> MSFT
        """

        if "Ticker Symbol" not in self.operations.columns:
            raise ValueError("'Ticker Symbol' column not found.")

        self.operations["Ticker Symbol"] = (
            self.operations["Ticker Symbol"]
            .astype(str)
            .str.split(".")
            .str[0]
            .str.strip()
        )

        return self.operations

    # ---------- SPLIT CFD AND STOCKS ----------
    def split_cfd_and_stocks(self):
        """
        Splits operations into CFD and regular stocks/ETFs.
        """

        if self.operations is None or self.operations.empty:
            raise ValueError("Operations dataframe is empty. Run normalize_closed_operations() first.")

        def is_cfd(ticker):
            ticker = str(ticker).strip()
            return bool(re.match(r"^[A-Z0-9]+$", ticker))

        # CFD operations
        self.cfd_operations = self.operations[
            self.operations["Ticker Symbol"].apply(is_cfd)
        ].copy()

        # Stock / ETF operations
        self.operations = self.operations[
            ~self.operations["Ticker Symbol"].apply(is_cfd)
        ].copy()

        return self.cfd_operations, self.operations

    # ---------- ADD DEPOSIT ----------
    def add_deposit(self, date=None):
        if date is None:
            date = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M")

        new_row = {
            "Type": "Deposit",
            "Date": date,
            "Value": self.header.get("Equity", 0),
            "Transaction Currency": self.account_currency,
            "Currency Gross Amount": self.account_currency,
            "Cash Account": self.account_currency,
            "Securities Account": f"XTB {self.account_currency}",
        }

        self.operations = pd.concat([self.operations, pd.DataFrame([new_row])], ignore_index=True)


    def export_default_cash_operations(self):
        try:
            self.read_header()
            self.read_table(["ID", "Type", "Time", "Comment", "Symbol", "Amount"])
            self.normalize_operations_history()
            self.strip_ticker_suffix()
            self.operations = self.operations[["Ticker Symbol", "Type", "Shares", "Date", "Value", "Securities Account", "Note"]]
            return self.operations
        except Exception as e:
            logging.exception(e)
            print(e)
            return pd.DataFrame()

    def export_open_operations(self):
        try:
            self.read_header()
            self.read_table(["Position", "Symbol", "Type", "Volume", "Open time", "Open price", "Market price", "Purchase value", "SL", "TP", "Margin", "Commission", "Swap", "Rollover", "Gross P/L", "Comment"])
            self.normalize_open_operations()
            self.strip_ticker_suffix()
            self.operations = self.operations[["Ticker Symbol", "Type", "Shares", "Date", "Value", "Securities Account", "Note"]]
            return self.operations
        except Exception as e:
            logging.exception(e)
            print(e)
            return pd.DataFrame()

    def export_closed_operations(self):
        try:
            self.read_header()
            self.read_table(["Position", "Symbol", "Type", "Volume", "Open time", "Open price", "Close time", "Close price", "Open origin", "Close origin", "Purchase value", "Sale value", "SL", "TP", "Margin", "Commission", "Swap", "Rollover", "Gross P/L", "Comment"])
            self.normalize_closed_operations()
            self.strip_ticker_suffix()
            self.operations = self.operations[["Ticker Symbol", "Type", "Shares", "Date", "Value", "Securities Account", "Note"]]
            return self.operations
        except Exception as e:
            logging.exception(e)
            print(e)
            return pd.DataFrame()

    def export_simplified_deposit_of_operation(self):
        try:
            self.read_header()
            self.add_deposit()
            return self.operations
        except Exception as e:
            logging.exception(e)
            print(e)
            return pd.DataFrame()


if __name__ == "__main__":
    logging.basicConfig(level=logging.NOTSET, filename="log.log", filemode="w", format="%(asctime)s - %(lineno)d - %(levelname)s - %(message)s")

    path = r""

    # ---------- CASH OPERATIONS HISTORY ----------
    cash_operations_history = CashOperationXLSXReader(path, sheet_index=3)
    cash_operations_history.export_default_cash_operations()
    #print(cash_operations_history.operations.tail())
    #cash_operations_history.operations.to_excel("cash_operations.xlsx", index=False)
    #cash_operations_history.operations.to_csv("cash_operations.csv", index=False, sep=',')

    # ---------- OPEN OPERATIONS ----------
    cash_open_operations = CashOperationXLSXReader(path, sheet_index=1)
    cash_open_operations.export_open_operations()
    #print(cash_open_operations.operations.tail())
    #cash_open_operations.operations.to_excel("cash_open_operations.xlsx", index=False)
    #cash_open_operations.operations.to_csv("cash_open_operations.csv", index=False, sep=',')

    # ---------- CLOSED OPERATIONS ----------
    cash_closed_operations = CashOperationXLSXReader(path, sheet_index=0)
    cash_closed_operations.export_closed_operations()
    #print(cash_closed_operations.operations.tail())
    #cash_closed_operations.operations.to_excel("cash_operations.xlsx", index=False)
    #cash_closed_operations.operations.to_csv("cash_closed_operations.csv", index=False, sep=',')

    # ---------- DEPOSIT ----------
    cash_deposit = CashOperationXLSXReader(path, sheet_index=3)
    cash_deposit.export_simplified_deposit_of_operation()
    #print(cash_deposit.operations.tail())
    #cash_deposit.operations.to_excel("cash_deposit.xlsx", index=False)
    #cash_deposit.operations.to_csv("cash_deposit.csv", index=False, sep=',
