import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import yfinance as yf
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime, timedelta
from tkcalendar import DateEntry
import csv
import threading
import queue

# Function to load tickers from a CSV file
def load_tickers(filename):
    tickers = []
    try:
        with open(filename, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)  # Skip header
            for row in reader:
                if len(row) >= 2:
                    ticker, name = row[0].strip().upper(), row[1].strip()
                    tickers.append(f"{ticker} - {name}")
    except FileNotFoundError:
        messagebox.showerror("File Error", f"'{filename}' file not found in the current directory.")
    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while loading tickers: {e}")
    return tickers

# Load tickers from 'tickers.csv'
TICKER_LIST = load_tickers('tickers.csv')  # Ensure 'tickers.csv' exists in the same directory

class AutocompleteCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list, key=lambda x: x.upper())
        self['values'] = self._completion_list

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._completion_list = []
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def autocomplete(self, delta=0):
        if delta:
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())
        
        # Get matching entries
        _hits = [element for element in self._completion_list if element.upper().startswith(self.get().upper())]
        
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits

        if _hits:
            self.delete(0, tk.END)
            self.insert(0, _hits[self._hit_index])
            self.select_range(self.position, tk.END)

    def handle_keyrelease(self, event):
        if event.keysym in ("BackSpace", "Left", "Right", "Up", "Down", "Return", "Escape"):
            return
        self.autocomplete()

class StockApp:
    def __init__(self, master):
        self.master = master
        master.title("Stock Market Analyzer")

        # Queue for threading
        self.queue = queue.Queue()

        # Create Widgets
        self.create_widgets()

        # Set default dates (last 30 days)
        today = datetime.today().date()
        thirty_days_ago = today - timedelta(days=30)
        self.start_date_entry.set_date(thirty_days_ago)
        self.end_date_entry.set_date(today)

        # Start queue processing
        self.master.after(100, self.process_queue)

    def create_widgets(self):
        # Frame for Stock Ticker
        ticker_frame = tk.Frame(self.master)
        ticker_frame.pack(pady=5)

        tk.Label(ticker_frame, text="Select Stock Ticker:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.ticker_entry = AutocompleteCombobox(ticker_frame, width=40)
        self.ticker_entry.set_completion_list(TICKER_LIST)
        self.ticker_entry.grid(row=0, column=1, padx=5, pady=5)

        # Frame for Start Date
        start_date_frame = tk.Frame(self.master)
        start_date_frame.pack(pady=5)

        tk.Label(start_date_frame, text="Start Date:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.start_date_entry = DateEntry(start_date_frame, width=18, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Frame for End Date
        end_date_frame = tk.Frame(self.master)
        end_date_frame.pack(pady=5)

        tk.Label(end_date_frame, text="End Date:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.end_date_entry = DateEntry(end_date_frame, width=18, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Get Data Button
        self.get_data_button = tk.Button(self.master, text="Get Data", command=self.get_stock_data)
        self.get_data_button.pack(pady=10)

        # Text Display for Stock Data
        self.text = tk.Text(self.master, height=10, width=80)
        self.text.pack(pady=10)

        # Plot Button
        self.plot_button = tk.Button(self.master, text="Plot Stock Data", command=self.plot_stock_data)
        self.plot_button.pack(pady=10)

    def get_stock_data(self):
        # Disable the Get Data button to prevent multiple clicks
        self.get_data_button.config(state=tk.DISABLED)
        threading.Thread(target=self.fetch_data).start()

    def fetch_data(self):
        input_text = self.ticker_entry.get().strip()
        if ' - ' in input_text:
            ticker, _ = input_text.split(' - ', 1)
        else:
            ticker = input_text.upper()

        start_date_str = self.start_date_entry.get_date().strftime('%Y-%m-%d')  # Get date from DateEntry
        end_date_str = self.end_date_entry.get_date().strftime('%Y-%m-%d')      # Get date from DateEntry

        # Validate Ticker Symbol
        if not ticker:
            self.queue.put(('error', "Please select or enter a stock ticker symbol."))
            self.queue.put(('enable_button', None))
            return

        # Ensure Start Date is Before End Date
        if start_date_str > end_date_str:
            self.queue.put(('error', "Start Date must be earlier than or equal to End Date."))
            self.queue.put(('enable_button', None))
            return

        # Fetch Stock Data
        try:
            stock = yf.Ticker(ticker)
            hist = stock.history(start=start_date_str, end=end_date_str)

            if hist.empty:
                self.queue.put(('error', f"No data found for {ticker} in the specified date range."))
                self.queue.put(('enable_button', None))
                return

            # Fetch Stock Information
            info = stock.info
            stock_name = info.get('longName', 'N/A')  # Get the long name, default to 'N/A' if not available

            # Get the Latest Available Data in the Specified Range
            latest = hist.iloc[-1]

            # Format Numerical Values
            formatted_open = f"{latest['Open']:,.2f}"
            formatted_high = f"{latest['High']:,.2f}"
            formatted_low = f"{latest['Low']:,.2f}"
            formatted_close = f"{latest['Close']:,.2f}"
            formatted_volume = f"{int(latest['Volume']):,}"  # Volume is an integer

            display_text = (
                f"Stock: {ticker} - {stock_name}\n"
                f"Date: {latest.name.date()}\n"
                f"Open: {formatted_open}\n"
                f"High: {formatted_high}\n"
                f"Low: {formatted_low}\n"
                f"Close: {formatted_close}\n"
                f"Volume: {formatted_volume}\n"
            )

            self.hist_data = hist
            self.queue.put(('success', display_text))

        except Exception as e:
            self.queue.put(('error', f"An error occurred while fetching data: {e}"))

        finally:
            self.queue.put(('enable_button', None))

    def process_queue(self):
        try:
            while True:
                msg_type, content = self.queue.get_nowait()
                if msg_type == 'success':
                    self.text.delete(1.0, tk.END)
                    self.text.insert(tk.END, content)
                elif msg_type == 'error':
                    messagebox.showerror("Error", content)
                elif msg_type == 'enable_button':
                    self.get_data_button.config(state=tk.NORMAL)
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_queue)

    def plot_stock_data(self):
        if not hasattr(self, 'hist_data'):
            messagebox.showerror("Error", "No data to plot. Please fetch stock data first.")
            return

        try:
            dates = self.hist_data.index
            closing_prices = self.hist_data['Close']

            # Create the Plot
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.plot(dates, closing_prices, marker='o', linestyle='-', label='Closing Price')
            ax.set_xlabel('Date')
            ax.set_ylabel('Price (USD)')
            ticker_display = self.ticker_entry.get().split(' - ')[0].upper()
            ax.set_title(f'Closing Prices for {ticker_display}')
            ax.legend()
            ax.grid(True)

            # Format the x-axis for dates
            fig.autofmt_xdate()

            # Clear previous plots
            for widget in self.master.winfo_children():
                if isinstance(widget, FigureCanvasTkAgg):
                    widget.get_tk_widget().destroy()

            # Embed the Plot in Tkinter
            canvas = FigureCanvasTkAgg(fig, master=self.master)
            canvas.draw()
            canvas.get_tk_widget().pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while plotting data: {e}")

def main():
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
