import os
from tkinter import messagebox, ttk
import requests
from bs4 import BeautifulSoup
import re
import random
import pandas as pd
import logging
import warnings
from logging import config
import sys
import time
import tkinter as tk
from tkinter import filedialog
import openpyxl
from webdriver_manager.chrome import ChromeDriverManager
import undetected_chromedriver as uc
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from contextlib import suppress

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

MAX_REQUEST_ATTEMPTS = 3

class CustomChrome(uc.Chrome):

    def __init__(self) -> None:
        super().__init__(
            driver_executable_path=ChromeDriverManager().install())
    
    def get(self, url: str) -> None:
        attempts_remaining = MAX_REQUEST_ATTEMPTS
        while attempts_remaining:
            with suppress(Exception):
                super().get(url)
                if "/try" not in self.current_url:
                    return
            attempts_remaining -= 1
        raise Exception(
            f"Failed to load the page even after {MAX_REQUEST_ATTEMPTS} "
            "attempts.")

# driver = CustomChrome()
driver = webdriver.Chrome()

warnings.filterwarnings("ignore")

config.dictConfig(
    {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {
            "default": {
                "format": "[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s"
            },
            "slack_format": {
                "format": "`[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d]` %(message)s"
            },
        },
        "handlers": {
            "file": {
                "class": "logging.FileHandler",
                "formatter": "default",
                "filename": "logs.log",
            },
        },
        "loggers": {
            "root": {
                "level": logging.INFO,
                "handlers": ["file"],
                "propagate": False,
            },
        },
    }
)
log = logging.getLogger("root")


def login(email, password):
    log.info("Signing in.")
    driver.get("https://kagi.com/signin")
    email_field = driver.find_element(By.ID, "signInEmailBox")
    email_field.send_keys(email)
    password_field = driver.find_element(By.ID, "signInPassBox")
    password_field.send_keys(password)

    # Find the sign-in button and click it
    sign_in_button = driver.find_element(By.XPATH, """//*[@id="signInForm"]/button""")
    sign_in_button.click()
    log.info("Signed in.")


def collect_query_list(df):
    queries = []
    for index, row in df.iterrows():
        name = row[0]
        unique_number = str(row[1]).zfill(8)
        queries.append(f"{name} {unique_number}")
    return queries


def get_website_text(query):
    log.info(f"Searching: {query}")
    url = f"https://kagi.com/search?q={query}"
    driver.get(url)
    time.sleep(2)
    # wait = WebDriverWait(driver, 10)
    # element = wait.until(EC.presence_of_element_located((By.XPATH, """//*[@id="load_more_results"]""")))
    soup = BeautifulSoup(driver.page_source, "lxml")
    texts = soup.find_all('div', class_='__sri-body')
    texts = [text.get_text() for text in texts]
    return texts


def extract_emails_manager(texts, name, number):
    emails = extract_emails(texts, name, number)
    if not emails:
        split_query = name.split(" ")
        if len(split_query) >= 3:
            emails = extract_emails(texts, " ".join([split_query[0], split_query[-1]]), number)
            if not emails:
                split_query.pop(0)
                emails = extract_emails(texts, split_query[-1], number)
        elif len(split_query) == 2:
            emails = extract_emails(texts, split_query[-1], number)
    return emails


def extract_emails(texts, name, number):
    email_list = []
    # Convert name and number to lowercase for case-insensitive matching
    name = name.lower()
    # Construct a case-insensitive regular expression pattern to match email addresses
    email_pattern = r"\b[A-Za-z0-9._%+-]+@(?!.*\.edu)[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    
    for text in texts:
        text = text.lower()
        if name in text and number in text:
            emails = re.findall(email_pattern, text)
            if emails:
                # Add the email address to the list
                email_list.extend(emails)
    # Log the found email IDs
    logging.info(f"Found email IDs: {list(set(email_list))}")
    # Return a list of unique email addresses
    return list(set(email_list))


def save_to_excel(data, excel_file_path="output.xlsx"):
    rows = []
    for name, unique_number, emails in data:
        skip_header = False
        for email in emails:
            rows.append(
                [
                    name if not skip_header else "",
                    unique_number if not skip_header else "",
                    email,
                ]
            )
            skip_header = True
    df = pd.DataFrame(rows)
    df.to_excel(excel_file_path, index=False)


def get_final_email_ids(query, name, number):
    x = get_website_text(query)
    emails = extract_emails_manager(x, name, number)
    if not emails:
        split_query = query.split("+")
        if len(split_query) >= 4:
            split_query = [split_query[0], split_query[-2], split_query[-1]]
            x = get_website_text("+".join(split_query))
            emails = extract_emails_manager(x, name, number)
            if not emails:
                split_query.pop(0)
                x = get_website_text("+".join(split_query))
                emails = extract_emails_manager(x, name, number)
        elif len(split_query) == 3:
            split_query.pop(0)
            x = get_website_text("+".join(split_query))
            emails = extract_emails_manager(x, name, number)
    return emails


def process_row(row, result_excel_file_path):
    query = get_query(row["NAME"], row["NUMBER"])
    emails = get_final_email_ids(query, row["NAME"], query.split("+")[-1])

    df = pd.DataFrame(
        {"NAME": [row["NAME"]], "NUMBER": [row["NUMBER"]], "EMAIL": [emails]}
    )
    if os.path.exists(result_excel_file_path):
        existing_df = pd.read_excel(
            result_excel_file_path, names=["NAME", "NUMBER", "EMAIL"], engine="openpyxl"
        )
        existing_df = pd.concat([existing_df, df], ignore_index=True)
        df = existing_df
    else:
        with open(result_excel_file_path, "w"):
            pass

    df = df.explode("EMAIL", ignore_index=True)
    duplicated_rows = df.duplicated(subset=["NAME", "NUMBER"])
    df.loc[duplicated_rows, ["NAME", "NUMBER"]] = ""
    log.info(f"Saved to excel: {result_excel_file_path}")

    def show_try_again_popup():
        result = messagebox.askretrycancel("Error", "Updating excel could not be possible. Please close the file if you are viewing")
        return result

    while True:
        try:
            df.to_excel(result_excel_file_path, index=False)
            break
        except:
            if not show_try_again_popup():
                continue


def get_query(name, number):
    return f"{name} {str(number).zfill(8)}".replace(" ", "+")


def main():
    root = tk.Tk()
    root.geometry("800x800")
    root.title("KAGI Email Scraper")
    output_file_name = 'email_list.xlsx'
    data = ""
    result_excel_file_path = ""
    username_var = tk.StringVar()
    password_var = tk.StringVar()

    def choose_source_file_path():
        nonlocal data
        data = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    def choose_save_path():
        nonlocal result_excel_file_path
        result_excel_file_path = filedialog.askdirectory()
        result_excel_file_path = os.path.join(result_excel_file_path, output_file_name)
        

    def submit():
        nonlocal root
        if data and result_excel_file_path:
            root.destroy()
            pass
        else:
            messagebox.showerror("Error", "Please choose both source and save paths before submitting.")

    label_source = tk.Label(root, text="Please choose your source excel file:")
    label_source.pack(pady=10)

    choose_source_button = tk.Button(root, text="Choose your excel sheet", command=choose_source_file_path)
    choose_source_button.pack(pady=5)

    label_save = tk.Label(root, text="Please choose the folder to save the result Excel file:")
    label_save.pack(pady=10)

    choose_path_button = tk.Button(root, text="Choose Save Path", command=choose_save_path)
    choose_path_button.pack(pady=5)

    label_username = tk.Label(root, text="Enter your KAGI username:")
    label_username.pack(pady=5)

    username_entry = tk.Entry(root, textvariable=username_var)
    username_entry.pack(pady=5)

    label_password = tk.Label(root, text="Enter your KAGI password:")
    label_password.pack(pady=5)

    password_entry = tk.Entry(root, textvariable=password_var)
    password_entry.pack(pady=5)

    submit_button = tk.Button(root, text="Submit", command=submit)
    submit_button.pack(pady=20)

    root.mainloop()

    if os.path.exists(result_excel_file_path):
        base_path, extension = os.path.splitext(result_excel_file_path)
        count = 1
        result_excel_file_path = f"{base_path}({count}){extension}"

        while os.path.exists(result_excel_file_path):
            count += 1
            result_excel_file_path = f"{base_path}({count}){extension}"

    df = pd.read_excel(data, header=None, engine="openpyxl")
    df.columns = ["NAME", "NUMBER"]

    progress_window = tk.Tk()
    progress_window.title("Progress: KAGI Email Scraper")

    progress_frame = ttk.Frame(progress_window)
    progress_frame.pack()

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
    progress_bar.grid(row=0, column=0, pady=5)

    total_rows = len(df)
    login(username_var.get(), password_var.get())
    for index, row in df.iterrows():
        process_row(row, result_excel_file_path)
        progress_bar["value"] = (index + 1) * 100 / total_rows
        progress_bar.update()

    progress_window.destroy()

if __name__ == "__main__":
    import traceback
    try:
        main()
    except Exception as e:
        log.error(traceback.format_exc())