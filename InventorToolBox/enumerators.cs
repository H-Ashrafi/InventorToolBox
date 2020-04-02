namespace InventorToolBox
{
    public enum KMainPlane
    {
        xy = 3, xz = 2, yz = 1
    }
    public enum kDocumnetProperty
    {
        //Design Tracking Properties
        Authority, CatalogWebLink, Categories, CheckedBy, Cost, CostCenter, CreationTime,
        DateChecked, DeferUpdates, Description, DesignStatus, Designer, DocumentSubType, DocumentSubTypeName,
        Engineer, EngrApprovedBy, EngrDateApproved, ExternalPropertyRevisionId, Language, Manufacturer, Material,
        MfgApprovedBy, MfgDateApproved, ParameterizedTemplate, PartIcon, PartNumber, PartPropertyRevisionId, Project, ProxyRefreshDate, SizeDesignation, Standard,
        StandardRevision, StandardsOrganization, StockNumber, TemplateRow, UserStatus, Vendor, WeldMaterial,
        //Inventor Document Summary Information
        Category, Company, Manager,
        //Inventor Summary Information
        Author, Comments, Keywords, LastSavedBy, Thumbnail, RevisionNumber, Subject, Title
    }
    public enum kManagerTypes
    {
        Document = 0,
        Part = 1,
        Assembly = 2,
        Drawing = 3,
        Presentation = 4,
        iProperties=5,
        BOM=6
    }
}
