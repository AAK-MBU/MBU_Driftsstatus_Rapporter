"""
This script contains the RPAReportGenerator class, which generates and sends RPA (Robotic Process Automation) reports via email.
The reports are fetched from an SQL Server database, processed, and presented in HTML format. The report includes information
about missed runs, failed processes, overdue processes, and process status.
"""
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pyodbc
from jinja2 import Template
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


class ReportGenerator:
    """
    A class to generate and send reports. The class retrieves data from an
    SQL Server database, processes it, and generates reports in HTML format
    that can be sent via email.
    """
    def __init__(self, orchestrator_connection: OrchestratorConnection, email_settings):
        """
        Initialize the RPAReportGenerator class with the orchestrator connection and email settings.

        :param orchestrator_connection: A connection object for interacting with the RPA orchestrator's database.
        :param email_settings: A dictionary containing SMTP email settings such as server, port, and credentials.
        """
        self.orchestrator_connection = orchestrator_connection
        self.email_settings = email_settings

    def get_db_connection(self):
        """Connect to the SQL Server database."""
        conn = pyodbc.connect(self.orchestrator_connection.get_constant('DbConnectionString').value)
        return conn

    def fetch_data(self, query):
        """Fetch data from the database using pyodbc and return as a list of tuples."""
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [column[0] for column in cursor.description]
        results = cursor.fetchall()
        conn.close()
        data = [dict(zip(columns, row)) for row in results]
        return data

    def missed_runs_report(self):
        """
        Generate a report of missed runs by comparing the `last_run` and `next_run` fields.

        :return: A list of dictionaries containing the report data for missed runs.
        """
        query = """
            SELECT t.trigger_name, t.process_name, t.last_run, st.next_run
            FROM [RPA].[dbo].[Triggers] t
            JOIN [RPA].[dbo].[Scheduled_Triggers] st ON t.id = st.id
            WHERE t.last_run < st.next_run
              AND GETDATE() > st.next_run
            ORDER BY trigger_name
        """
        return self.fetch_data(query)

    def process_failure_report(self):
        """
        Generate a report of processes that failed in the last 7 days.

        :return: A list of dictionaries containing the report data for failed processes.
        """
        query = """
            SELECT t.trigger_name, t.process_name, t.last_run, t.process_status
            FROM [RPA].[dbo].[Triggers] t
            WHERE t.process_status = 'Failed'
              AND t.last_run > DATEADD(day, -7, GETDATE())
            ORDER BY trigger_name
        """
        return self.fetch_data(query)

    def process_status_report(self):
        """
        Generate a status report for all processes with their latest run information.

        :return: A list of dictionaries containing process status data.
        """
        query = """
            SELECT t.trigger_name, t.process_status, st.next_run, t.last_run
            FROM [RPA].[dbo].[Triggers] t
            JOIN [RPA].[dbo].[Scheduled_Triggers] st ON t.id = st.id
            WHERE t.last_run IS NOT NULL
            ORDER BY trigger_name
        """
        return self.fetch_data(query)

    def overdue_processes_report(self):
        """
        Generate a report of processes that have not run by their scheduled next run time.

        :return: A list of dictionaries containing the report data for overdue processes.
        """
        query = """
            SELECT t.trigger_name, t.process_name, st.next_run, t.last_run
            FROM [RPA].[dbo].[Triggers] t
            JOIN [RPA].[dbo].[Scheduled_Triggers] st ON t.id = st.id
            WHERE st.next_run < GETDATE()
            ORDER BY trigger_name
        """
        return self.fetch_data(query)

    def generate_html_report(self):
        """
        Generate the HTML report for all process data, including missed runs, failed processes,
        overdue processes, and process statuses.

        :return: A tuple containing the generated HTML content and an alert flag indicating
                 whether there are any missed or overdue processes.
        """
        missed_runs = self.missed_runs_report()
        overdue_processes = self.overdue_processes_report()
        process_failures = self.process_failure_report()
        process_status = self.process_status_report()

        alert_flag = bool(missed_runs or overdue_processes)

        html_template = """
        <html>
        <head>
        <style>
            table, th, td {
                border: 1px solid black;
                border-collapse: collapse;
                padding: 8px;
                font-size: 14px;  /* Mindre fontstørrelse for tabeller */
            }
            h3 {
                font-size: 16px;  /* Mindre fontstørrelse for overskrifter */
            }
            body {
                font-size: 12px;  /* Generel fontstørrelse for hele dokumentet */
            }
        </style>
        </head>
        <body>
            <h3>Manglende kørsler</h3>
            {{ missed_runs_table | safe }}
        <br/><br/>
            <h3>Fejlede processer (Sidste 7 Dage)</h3>
            {{ process_failures_table | safe }}
        <br/><br/>
            <h3>Forsinkede processer</h3>
            {{ overdue_processes_table | safe }}
        <br/><br/>
            <h3>Processtatus</h3>
            {{ process_status_table | safe }}
        </body>
        </html>
        """

        missed_runs_table = self.convert_to_html_table(missed_runs)
        overdue_processes_table = self.convert_to_html_table(overdue_processes)
        process_failures_table = self.convert_to_html_table(process_failures)
        process_status_table = self.convert_to_html_table(process_status)

        template = Template(html_template)
        html_content = template.render(
            missed_runs_table=missed_runs_table,
            overdue_processes_table=overdue_processes_table,
            process_failures_table=process_failures_table,
            process_status_table=process_status_table,
        )

        return html_content, alert_flag

    def convert_to_html_table(self, data):
        """
        Convert a list of dictionaries into an HTML table format.

        :param data: A list of dictionaries where each dictionary represents a row of data.
        :return: A string representing the data as an HTML table.
        """
        if not data:
            return "<p>Ingen data tilgængelig.</p>"

        headers = data[0].keys()
        header_html = ''.join(f'<th>{header}</th>' for header in headers)

        rows_html = ''.join(
            '<tr>' + ''.join(f'<td>{value}</td>' for value in row.values()) + '</tr>' 
            for row in data
        )

        html_table = f'<table><thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table>'
        return html_table

    def send_email(self, html_content):
        """
        Send the generated HTML report via email. If there are missed or overdue processes,
        mark the email as high priority.

        :param html_content: The HTML content to be sent in the email body.
        """
        html_content, alert_flag = self.generate_html_report()
        subject = 'RPA - Driftsrapport'

        msg = MIMEMultipart()

        if alert_flag:
            subject += ' - ⚠️ Advarsel. Handling påkrævet ⚠️'
            msg['X-Priority'] = '1'
            msg['X-MSMail-Priority'] = 'High'
            msg['Importance'] = 'High'

        msg['From'] = self.email_settings['from_email']
        msg['To'] = self.email_settings['to_email']
        msg['Subject'] = subject

        msg.attach(MIMEText(html_content, 'html'))

        with smtplib.SMTP(self.email_settings['smtp_server'], self.email_settings['smtp_port']) as server:
            server.starttls()
            server.send_message(msg)
            print(f"Email sent to {self.email_settings['to_email']}.")
