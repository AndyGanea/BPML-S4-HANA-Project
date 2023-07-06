import os
import logging

logger = logging.getLogger(__name__)


def _detect_region(http_client=None):
    region = os.environ.get("REGION_NAME", "").replace(" ", "").lower()  # e.g. westus2
    if region:
        return region
    if http_client:
        return _detect_region_of_azure_vm(http_client)  # It could hang for minutes
    return None


def _detect_region_of_azure_vm(http_client):
    url = (
        "http://169.254.169.254/metadata/instance"

        # Utilize the "route parameters" feature to obtain region as a string
        # https://docs.microsoft.com/en-us/azure/virtual-machines/windows/instance-metadata-service?tabs=linux#route-parameters
        "/compute/location?format=text"

        # Location info is available since API version 2017-04-02
        # https://docs.microsoft.com/en-us/azure/virtual-machines/windows/instance-metadata-service?tabs=linux#response-1
        "&api-version=2021-01-01"
        )
    logger.info(
        "Connecting to IMDS {}. "
        "It may take a while if you are running outside of Azure. "
        "You should consider opting in/out region behavior on-demand, "
        'by loading a boolean flag "is_deployed_in_azure" '
        'from your per-deployment config and then do '
        '"app = ConfidentialClientApplication(..., '
        'azure_region=is_deployed_in_azure)"'.format(url))
    try:
        # https://docs.microsoft.com/en-us/azure/virtual-machines/windows/instance-metadata-service?tabs=linux#instance-metadata
        resp = http_client.get(url, headers={"Metadata": "true"})
    except:
        logger.info(
            "IMDS {} unavailable. Perhaps not running in Azure VM?".format(url))
        return None
    else:
        return resp.text.strip()

