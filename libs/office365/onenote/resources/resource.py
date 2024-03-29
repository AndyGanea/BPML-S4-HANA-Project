from office365.onenote.entity_base_model import OnenoteEntityBaseModel
from office365.runtime.client_result import ClientResult
from office365.runtime.queries.function import FunctionQuery

class OnenoteResource(OnenoteEntityBaseModel):
    """An image or other file resource on a OneNote page."""

    def get_content(self):
        """Retrieve the binary data of a file or image resource object."""
        return_type = ClientResult(self.context)
        qry = FunctionQuery(self, "content", None, return_type)
        self.context.add_query(qry)
        return return_type

    @property
    def content_url(self):
        """The URL for downloading the content

        :rtype: str or None
        """
        return self.properties.get("contentUrl", None)
