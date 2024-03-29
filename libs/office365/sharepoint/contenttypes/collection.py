from office365.runtime.queries.create_entity import CreateEntityQuery
from office365.runtime.queries.service_operation import ServiceOperationQuery
from office365.runtime.paths.service_operation import ServiceOperationPath
from office365.sharepoint.base_entity_collection import BaseEntityCollection
from office365.sharepoint.contenttypes.content_type import ContentType
from office365.sharepoint.contenttypes.entity_data import ContentTypeEntityData


class ContentTypeCollection(BaseEntityCollection):
    """Content Type resource collection"""

    def __init__(self, context, resource_path=None, parent=None):
        super(ContentTypeCollection, self).__init__(context, ContentType, resource_path, parent)

    def get_by_name(self, name):
        """
        Returns the content type with the given name from the collection.

        :param str name: Content type name
        """
        return_type = ContentType(self.context)
        self.add_child(return_type)

        def _after_get_by_name(col):
            """
            :type col: ContentTypeCollection
            """
            if len(col) != 1:
                message = "Content type not found or ambiguous match found for name: {0}".format(name)
                raise ValueError(message)
            return_type.set_property("StringId", col[0].get_property("StringId"))

        self.filter("Name eq '{0}'".format(name))
        self.context.load(self, after_loaded=_after_get_by_name)
        return return_type

    def get_by_id(self, content_type_id):
        """
        Returns the content type with the given identifier from the collection.
        If a content type with the given identifier is not found in the collection, the server MUST return null.

        :param str content_type_id: A hexadecimal value representing the identifier of a content type.
        """
        return ContentType(self.context, ServiceOperationPath("GetById", [content_type_id], self.resource_path))

    def add(self, content_type_info):
        """Adds a new content type to the collection and returns a reference to the added SP.ContentType.

        :param ContentTypeCreationInformation content_type_info: Specifies properties that is to be used to
            construct the new content type.
        """
        return_type = ContentType(self.context)
        self.add_child(return_type)
        params = content_type_info.to_json()
        for k, v in params.items():
            if k == "Id":
                return_type.set_property(k, {"StringValue": v}, True)
            else:
                return_type.set_property(k, v, True)
        qry = CreateEntityQuery(self, return_type, return_type)
        self.context.add_query(qry)
        return return_type

    def create(self, name, description=None, group=None, parent_content_type=None):
        """
        Creates a new content type to the collection and returns a reference to the added SP.ContentType.

        :param str name:  Specifies the name
        :param str description: Specifies the description
        :param str group: Specifies the group of the content type
        :param str or ContentType parent_content_type: Specifies the parent content type (string identifier or object)
        """
        def _create_query(parent_content_type_id):
            """
            :type parent_content_type_id: str
            """
            parameters = ContentTypeEntityData(name, description, group, parent_content_type_id)
            return ServiceOperationQuery(self, "Create", None, parameters, "parameters", return_type)

        return_type = ContentType(self.context)
        self.add_child(return_type)
        if isinstance(parent_content_type, ContentType):
            def _ct_loaded():
                next_qry = _create_query(parent_content_type.string_id)
                self.context.add_query(next_qry)
            parent_content_type.ensure_property("StringId", _ct_loaded)
        else:
            qry = _create_query(parent_content_type)
            self.context.add_query(qry)
        return return_type

    def add_available_content_type(self, content_type_id):
        """Adds the specified content type to the content type collection.

        :param str content_type_id: Specifies the identifier of the content type to be added to the content type
            collection. It MUST exist in the web's available content types.
        """
        return_type = ContentType(self.context)
        self.add_child(return_type)
        qry = ServiceOperationQuery(self, "AddAvailableContentType", [content_type_id], None, None, return_type)
        self.context.add_query(qry)
        return return_type
