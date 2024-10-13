# SLB Industrial Training and Safety Email Reminder System

This project automates the process of sending reminders to employees about the expiration status of their industrial training certificates and safety instructions. It categorizes employees based on the expiration timeline and sends timely email notifications to both employees and their managers.

## Features

- **Automated Email Notifications:**  
  Sends emails to employees whose training or safety certifications are nearing expiration or have already expired.
  
- **Expiration Categories:**  
  Employees are notified based on these timeframes:
  - **Expired** (Past due)
  - **Expiring within 7 days**
  - **Expiring within 14 days**
  - **Expiring within 30 days**

- **Excel Data Integration:**  
  The program reads data from multiple Excel files to track different categories:
  - **D&M Safety Instructions**  
  - **TLM Safety Instructions**  
  - **Industrial Training Requirements**  

- **Customizable Email Content:**  
  Emails are color-coded and tailored based on the expiration category for better visibility.

## How It Works

1. **Excel Data Processing:**  
   The program loads multiple Excel sheets containing training and safety instruction data for employees.  
2. **Categorization:**  
   Employees are categorized based on the number of days left before their certificates expire.  
3. **Email Notification:**  
   The system sends out HTML-formatted emails with the appropriate message and background color, notifying employees and their managers.  
4. **File Attachment:**  
   A weekly summary Excel file is attached to the main notification email.



