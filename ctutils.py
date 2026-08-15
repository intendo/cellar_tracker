"""A collection of utility functions."""

from logging import getLevelName, getLogger, Formatter, INFO
from os import getenv

# Electrostatics/apbs-aws/lambda_services/job_service/launcher/utils.py

def this_logger():
    """Get a singleton logger for all code.

    Returns:
        Logger: An all encompassing logger.
    """
    # Override Lambda's log format
    root_logger = getLogger()
    for handler in root_logger.handlers:
        handler.setFormatter(
            Formatter(
                "[%(levelname)s] "
                "[%(filename)s:%(lineno)s:%(funcName)s()] %(message)s"
            )
        )

    # Set log level for our context
    _this_logger = getLogger(__name__)
    _this_logger.setLevel(getenv("LOG_LEVEL", getLevelName(INFO)))
    return _this_logger


_LOGGER = this_logger()

# Example usage:
#    _LOGGER.debug("ITEM: %s", item)
