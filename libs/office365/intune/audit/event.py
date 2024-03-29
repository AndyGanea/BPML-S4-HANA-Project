from office365.entity import Entity
from office365.entity_collection import EntityCollection
from office365.runtime.client_result import ClientResult
from office365.runtime.queries.function import FunctionQuery
from office365.runtime.types.collections import StringCollection


class AuditEvent(Entity):
    """A class containing the properties for Audit Event."""


class AuditEventCollection(EntityCollection):

    def __init__(self, context, resource_path=None):
        super(AuditEventCollection, self).__init__(context, AuditEvent, resource_path)

    def get_audit_categories(self):
        return_type = ClientResult(self.context, StringCollection())
        qry = FunctionQuery(self, "getAuditCategories", None, return_type)
        self.context.add_query(qry)
        return return_type

    def get_audit_activity_types(self, category):
        """"""
        return_type = ClientResult(self.context, StringCollection())
        params = {"category": category}
        qry = FunctionQuery(self, "getAuditActivityTypes", params, return_type)
        self.context.add_query(qry)
        return return_type

