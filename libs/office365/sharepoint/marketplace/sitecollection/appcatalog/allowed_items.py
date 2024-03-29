from office365.sharepoint.base_entity_collection import BaseEntityCollection
from office365.sharepoint.marketplace.sitecollection.appcatalog.allowed_item import SiteCollectionAppCatalogAllowedItem


class SiteCollectionAppCatalogAllowedItems(BaseEntityCollection):

    def __init__(self, context, resource_path=None):
        super(SiteCollectionAppCatalogAllowedItems, self).__init__(context,
                                                                   SiteCollectionAppCatalogAllowedItem,resource_path)
