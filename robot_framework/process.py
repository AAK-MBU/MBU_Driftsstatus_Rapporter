"""This module contains the main process of the robot."""

import json
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from robot_framework import config
from robot_framework.sub_processes.report_handler import ReportGenerator


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    oc_args_json = json.loads(orchestrator_connection.process_arguments)

    email_settings = {
        'from_email': f'{oc_args_json["fromEmail"]}',
        'to_email': f'{oc_args_json["toEmail"]}',
        'smtp_server': f'{config.SMTP_SERVER}',
        'smtp_port': f'{config.SMTP_PORT}'
        }

    report_generator = ReportGenerator(
        orchestrator_connection=orchestrator_connection,
        email_settings=email_settings
    )

    html_report = report_generator.generate_html_report()
    report_generator.send_email(html_report)


if __name__ == "__main__":
    oc = OrchestratorConnection.create_connection_from_args()
    process(oc)
