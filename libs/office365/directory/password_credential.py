from office365.runtime.client_value import ClientValue


class PasswordCredential(ClientValue):
    """Represents a password credential associated with an application or a service principal.
    The passwordCredentials property of the application entity is a collection of passwordCredential objects."""

    def __init__(self, display_name=None, secret_text=None, key_id=None):
        """
        :param str display_name: Friendly name for the password. Optional.
        :param str secret_text: Read-only; Contains the strong passwords generated by Azure AD that are 16-64
            characters in length. The generated password value is only returned during the initial POST request to
            addPassword. There is no way to retrieve this password in the future.
        :param str key_id: The unique identifier for the password.
        """
        super(PasswordCredential, self).__init__()
        self.displayName = display_name
        self.secretText = secret_text
        self.keyId = key_id
