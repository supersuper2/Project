import tkinter as tk
from tkinter import messagebox
import json
import smtplib
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import time
from openpyxl import Workbook
import os

# Azure imports
from azure.identity import AzureCliCredential
from azure.mgmt.costmanagement import CostManagementClient
from azure.core.exceptions import HttpResponseError

def create_excel_report(rows, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Azure Cost Report"

    # Headers
    ws.append(["Date", "Service", "Cost (USD)"])

    total_cost = 0.0

    for row in rows:
        try:
            cost = float(row[0])
            raw_date = row[1]
            service = row[2]
            date = datetime.strptime(str(raw_date), "%Y%m%d").strftime("%Y-%m-%d")

            ws.append([date, service, cost])
            total_cost += cost
        except (ValueError, TypeError, IndexError):
            continue

    # Add total at the end
    ws.append(["", "TOTAL", total_cost])
    wb.save(file_path)
    return total_cost


def send_email_with_attachment(to_addresses, body, attachment_path):
    for to_address in to_addresses:
        msg = EmailMessage()
        msg.set_content(body)
        msg["Subject"] = "Azure Monthly Spend Report"
        msg["From"] = "s.simar123@gmail.com"
        msg["To"] = to_address

        # Attach Excel file
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
            msg.add_attachment(part.get_payload(decode=True), maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=os.path.basename(attachment_path))

        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login("s.simar123@gmail.com", "teis cxgy mspz wnhh")  # App password
                server.send_message(msg)
        except Exception as e:
            print(f"Failed to send to {to_address}: {e}")


def button1_action():
    try:
        credential = AzureCliCredential()
        client = CostManagementClient(credential)

        subscription_id = "f131ab3d-ecbf-4ce4-b2c2-c937dc44a331"
        scope = f"/subscriptions/{subscription_id}"

        # Load email recipients
        try:
            with open("email_mapping.json") as f:
                mapping = json.load(f)
                recipients = mapping.get("recipients", [])
        except Exception as e:
            messagebox.showerror("Email Mapping Error", str(e))
            return

        parameters = {
            "type": "ActualCost",
            "timeframe": "MonthToDate",
            "dataset": {
                "granularity": "Daily",
                "aggregation": {
                    "totalCost": {"name": "PreTaxCost", "function": "Sum"},
                },
                "grouping": [
                    {
                        "type": "Dimension",
                        "name": "ServiceName"
                    }
                ]
            }
        }

        result = client.query.usage(scope=scope, parameters=parameters)
        time.sleep(2)
        rows = result.rows

        if rows and all(isinstance(col, str) for col in rows[0]):
            rows = rows[1:]

        file_path = "azure_cost_report.xlsx"
        total_cost = create_excel_report(rows, file_path)

        summary = f"Azure Monthly Spend Report\n\nTotal Subscription Cost (MonthToDate): ${total_cost:.2f}\n\nSee attached Excel file for details."
        send_email_with_attachment(recipients, summary, file_path)

        messagebox.showinfo("Monthly Spend Report", "Report sent successfully!")

    except HttpResponseError as e:
        messagebox.showerror("Azure API Error", f"Error retrieving data:\n{e}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")


# Tkinter GUI setup
root = tk.Tk()
root.title("Azure Cost Project")
root.geometry("300x200")

button1 = tk.Button(root, text="Send Spend Report", command=button1_action)
button1.pack(pady=10)

root.mainloop()