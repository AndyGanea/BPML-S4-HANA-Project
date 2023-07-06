from office365.runtime.client_result import ClientResult
from office365.runtime.queries.service_operation import ServiceOperationQuery
from office365.sharepoint.activities.tracked_item_updates_request import TrackedItemUpdatesRequest
from office365.sharepoint.base_entity import BaseEntity


class TrackedItemService(BaseEntity):

    @staticmethod
    def get_tracked_item_updates_for_user(context):
        """
        :type context: office365.sharepoint.client_context.ClientContext
        """
        return_type = ClientResult(context)
        payload = {
            "request": TrackedItemUpdatesRequest()
        }
        qry = ServiceOperationQuery(TrackedItemService(context), "GetTrackedItemUpdatesForUser", None,
                                    payload, None, return_type, True)
        context.add_query(qry)
        return return_type

    @property
    def entity_type_name(self):
        return "Microsoft.SharePoint.Internal.TrackedItemService"
