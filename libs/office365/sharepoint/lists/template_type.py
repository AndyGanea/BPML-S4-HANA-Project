class ListTemplateType:
    """Specifies the type of a list template. A set of predefined values are specified in [MS-WSSTS] section 2.7."""

    def __init__(self):
        pass

    InvalidType = -1

    NoListTemplate = 0
    """Specifies no template applies."""

    GenericList = 100
    """Specifies a basic list"""

    DocumentLibrary = 101
    """Specifies a Document library"""

    Survey = 102
    """Specifies a survey list."""

    Links = 103

    Announcements = 104

    Contacts = 105
    """Specifies a contacts list."""

    Events = 106

    Tasks = 107

    DiscussionBoard = 108

    PictureLibrary = 109

    DataSources = 110

    WebTemplateCatalog = 111

    UserInformation = 112

    WebPartCatalog = 113

    ListTemplateCatalog = 114

    XMLForm = 115

    MasterPageCatalog = 116

    NoCodeWorkflows = 117

    WorkflowProcess = 118

    WebPageLibrary = 119

    CustomGrid = 120

    SolutionCatalog = 121

    NoCodePublic = 122

    ThemeCatalog = 123

    DesignCatalog = 124

    AppDataCatalog = 125

    DataConnectionLibrary = 130

    WorkflowHistory = 140

    GanttTasks = 150

    HelpLibrary = 151

    AccessRequest = 160

    TasksWithTimelineAndHierarchy = 171

    MaintenanceLogs = 175

    Meetings = 200

    Agenda = 201

    MeetingUser = 202

    Decision = 204

    MeetingObjective = 207

    TextBox = 210

    ThingsToBring = 211

    HomePageLibrary = 212

    Posts = 301

    Comments = 302

    Categories = 303

    Facility = 402

    Whereabouts = 403

    CallTrack = 404

    Circulation = 405

    Timecard = 420

    Holidays = 421

    IMEDic = 499
