import os  # Added for file path operations
from PIL import Image  # Added for image creation
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import yfinance as yf
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from tkcalendar import DateEntry
import pandas as pd  # Added pandas for Excel handling
import threading
import queue

# Define the path for the blank icon
ICON_PATH = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    if not os.path.exists(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    if not os.path.exists(path):
        size = (16, 16)  # Size of the icon
        image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
        image.save(path, format="ICO")

# Create the blank icon
create_blank_ico(ICON_PATH)

# Function to load tickers from an Excel file
def load_tickers(filename):
    tickers = []
    try:
        # Read the Excel file using pandas
        df = pd.read_excel(filename)
        
        # Check if the DataFrame has at least two columns
        if df.shape[1] < 2:
            raise ValueError("Excel file must contain at least two columns: Ticker and Name.")
        
        # Iterate over the rows and extract ticker and name
        for index, row in df.iterrows():
            ticker = str(row.iloc[0]).strip().upper()
            name = str(row.iloc[1]).strip()
            if ticker and name:
                tickers.append(f"{ticker} - {name}")
                
        if not tickers:
            raise ValueError("No valid ticker entries found in the Excel file.")
                
    except FileNotFoundError:
        messagebox.showerror("File Error", f"'{filename}' file not found in the current directory.")
    except ValueError as ve:
        messagebox.showerror("File Format Error", str(ve))
    except Exception as e:
        messagebox.showerror("File Error", f"An error occurred while loading tickers: {e}")
    return tickers

# Load tickers from 'tickers.xlsx'
TICKER_LIST = load_tickers('tickers.xlsx')  # Ensure 'tickers.xlsx' exists in the same directory

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
        master.title("Stock Market Monitor")
        
        # Set the window icon to the blank icon
        try:
            master.iconbitmap(ICON_PATH)
        except Exception as e:
            messagebox.showwarning("Icon Error", f"Failed to set window icon: {e}")

        # Initialize ttk.Style
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Use 'clam' theme for better customization

        # Define custom style for buttons
        self.style.configure("Custom.TButton",
                             background="#d0e8f1",
                             foreground="black",
                             borderwidth=1,
                             focusthickness=3,
                             focuscolor='none')

        # Define style map for hover (active) state
        self.style.map("Custom.TButton",
                       background=[('active', '#87CEFA')],
                       foreground=[('active', 'black')])

        # Define widget styles to avoid conflict with ttk.Style
        self.widget_style = {"background": "#f0f0f0", "foreground": "#333333", "font": ("Arial", 11)}

        # Initialize hist_data
        self.hist_data = None

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
        ticker_frame = tk.Frame(self.master, bg=self.widget_style["background"])
        ticker_frame.pack(pady=5, padx=10, fill='x')

        tk.Label(ticker_frame, text="Select Stock Ticker:", **self.widget_style).grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.ticker_entry = AutocompleteCombobox(ticker_frame, width=40)
        self.ticker_entry.set_completion_list(TICKER_LIST)
        self.ticker_entry.grid(row=0, column=1, padx=5, pady=5)

        # Frame for Start Date
        start_date_frame = tk.Frame(self.master, bg=self.widget_style["background"])
        start_date_frame.pack(pady=5, padx=10, fill='x')

        tk.Label(start_date_frame, text="Start Date:", **self.widget_style).grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.start_date_entry = DateEntry(start_date_frame, width=18, date_pattern='yyyy-mm-dd')
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Frame for End Date
        end_date_frame = tk.Frame(self.master, bg=self.widget_style["background"])
        end_date_frame.pack(pady=5, padx=10, fill='x')

        tk.Label(end_date_frame, text="End Date:", **self.widget_style).grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.end_date_entry = DateEntry(end_date_frame, width=18, date_pattern='yyyy-mm-dd')
        self.end_date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Get Data Button
        self.get_data_button = ttk.Button(self.master, text="Get Data", command=self.get_stock_data, style="Custom.TButton")
        self.get_data_button.pack(pady=10)

        # Text Display for Stock Data
        self.text = tk.Text(self.master, height=10, width=80, bg=self.widget_style["background"], fg=self.widget_style["foreground"], font=self.widget_style["font"], borderwidth=1, relief="solid")
        self.text.pack(pady=10, padx=10)

        # Plot Button
        self.plot_button = ttk.Button(self.master, text="Plot Stock Data", command=self.plot_stock_data, style="Custom.TButton")
        self.plot_button.pack(pady=10)

        # Calculate Percentage Difference Button
        self.calc_button = ttk.Button(self.master, text="Calculate % Difference", command=self.calculate_percentage_difference, style="Custom.TButton")
        self.calc_button.pack(pady=10)

        # Label to Display Percentage Difference
        self.percentage_label = tk.Label(self.master, text="Percentage Difference: N/A", bg=self.widget_style["background"], fg=self.widget_style["foreground"], font=("Arial", 12))
        self.percentage_label.pack(pady=5)

    def get_stock_data(self):
        # Disable the Get Data button to prevent multiple clicks
        self.get_data_button.config(state="disabled")  # Corrected state to "disabled"
        threading.Thread(target=self.fetch_data, daemon=True).start()

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

            # Instead of assigning self.hist_data here, pass it through the queue
            self.queue.put(('success', display_text, hist))
            self.queue.put(('reset_percentage', None))  # Reset percentage difference display

        except Exception as e:
            self.queue.put(('error', f"An error occurred while fetching data: {e}"))

        finally:
            self.queue.put(('enable_button', None))

    def process_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                msg_type = msg[0]
                if msg_type == 'success':
                    display_text, hist = msg[1], msg[2]
                    self.text.delete(1.0, tk.END)
                    self.text.insert(tk.END, display_text)
                    self.hist_data = hist  # Assign hist_data in the main thread
                elif msg_type == 'error':
                    content = msg[1]
                    messagebox.showerror("Error", content)
                elif msg_type == 'enable_button':
                    self.get_data_button.config(state="normal")  # Corrected state to "normal"
                elif msg_type == 'reset_percentage':
                    self.percentage_label.config(text="Percentage Difference: N/A")
        except queue.Empty:
            pass
        finally:
            self.master.after(100, self.process_queue)

    def plot_stock_data(self):
        if not hasattr(self, 'hist_data') or self.hist_data is None:
            messagebox.showerror("Error", "No data to plot. Please fetch stock data first.")
            return

        try:
            dates = self.hist_data.index
            closing_prices = self.hist_data['Close']

            # Create the Plot
            plt.figure(figsize=(10, 5))
            plt.plot(dates, closing_prices, marker='o', linestyle='-', label='Closing Price')
            ticker_display = self.ticker_entry.get().split(' - ')[0].upper()
            plt.xlabel('Date')
            plt.ylabel('Price (USD)')
            plt.title(f'Closing Prices for {ticker_display}')
            plt.legend()
            plt.grid(True)
            plt.xticks(rotation=45)
            plt.tight_layout()

            # Show the plot in a separate window
            plt.show()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while plotting data: {e}")

    def calculate_percentage_difference(self):
        if not hasattr(self, 'hist_data') or self.hist_data is None:
            messagebox.showerror("Error", "No data available. Please fetch stock data first.")
            return

        try:
            start_date = self.start_date_entry.get_date()
            end_date = self.end_date_entry.get_date()

            # Ensure that hist_data is sorted by date
            hist_sorted = self.hist_data.sort_index()

            # Find the closest available date on or after the start date
            available_start_dates = [date for date in hist_sorted.index if date.date() >= start_date]
            if not available_start_dates:
                messagebox.showerror("Error", "No trading data available on or after the start date.")
                return
            actual_start_date = available_start_dates[0].date()
            start_closing = hist_sorted.loc[available_start_dates[0]]['Close']

            # Find the closest available date on or before the end date
            available_end_dates = [date for date in reversed(hist_sorted.index) if date.date() <= end_date]
            if not available_end_dates:
                messagebox.showerror("Error", "No trading data available on or before the end date.")
                return
            actual_end_date = available_end_dates[0].date()
            end_closing = hist_sorted.loc[available_end_dates[0]]['Close']

            # Calculate percentage difference
            percentage_diff = ((end_closing - start_closing) / start_closing) * 100
            formatted_diff = f"{percentage_diff:.2f}%"

            self.percentage_label.config(
                text=f"Percentage Difference from {actual_start_date} to {actual_end_date}: {formatted_diff}"
            )

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while calculating percentage difference: {e}")

def main():
    root = tk.Tk()
    app = StockApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
