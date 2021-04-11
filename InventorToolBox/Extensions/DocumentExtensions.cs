using Inventor;
using System;

namespace InventorToolBox
{
    /// <summary>
    /// extension for <see cref="Document"/>
    /// </summary>
    public static class DocumentExtensions
    {
        /// <summary>
        /// gets the value of a property based on its property-set and name
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="propetySet"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        private static Property GetProperty(Document doc, string propetySet, string propertyName)
        {
            try
            {
                PropertySets propSets = doc.PropertySets;
                PropertySet propSet = propSets[propetySet];
                Property prop = propSet[propertyName];
                return prop;
            }
            catch
            {
                throw new Exception($"Could not find the {propertyName} property");
            }
        }

        /// <summary>
        /// get value of a <see cref="Property"/> in iProperty 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="propertyType">the type of property to retrive the value of</param>
        /// <returns></returns>
        public static Property GetProperty(this Document doc, kDocumnetProperty propertyType)
        {

            string trackSet = "Design Tracking Properties";
            string docSet = "Inventor Document Summary Information";
            string sumSet = "Inventor Summary Information";

            switch (propertyType)
            {
                case kDocumnetProperty.Authority:
                    return GetProperty(doc, trackSet, "Authority");
                case kDocumnetProperty.CatalogWebLink:
                    return GetProperty(doc, trackSet, "Catalog Web Link");
                case kDocumnetProperty.Categories:
                    return GetProperty(doc, trackSet, "Categories");
                case kDocumnetProperty.CheckedBy:
                    return GetProperty(doc, trackSet, "Checked By");
                case kDocumnetProperty.Cost:
                    return GetProperty(doc, trackSet, "Cost");
                case kDocumnetProperty.CostCenter:
                    return GetProperty(doc, trackSet, "Cost Center");
                case kDocumnetProperty.CreationTime:
                    return GetProperty(doc, trackSet, "Creation Time");
                case kDocumnetProperty.DateChecked:
                    return GetProperty(doc, trackSet, "Date Checked");
                case kDocumnetProperty.DeferUpdates:
                    return GetProperty(doc, trackSet, "Defer Updates");
                case kDocumnetProperty.Description:
                    return GetProperty(doc, trackSet, "Description");
                case kDocumnetProperty.DesignStatus:
                    return GetProperty(doc, trackSet, "Design Status");
                case kDocumnetProperty.Designer:
                    return GetProperty(doc, trackSet, "Designer");
                case kDocumnetProperty.DocumentSubType:
                    return GetProperty(doc, trackSet, "Document SubType");
                case kDocumnetProperty.DocumentSubTypeName:
                    return GetProperty(doc, trackSet, "Document SubType Name");
                case kDocumnetProperty.Engineer:
                    return GetProperty(doc, trackSet, "Engineer");
                case kDocumnetProperty.EngrApprovedBy:
                    return GetProperty(doc, trackSet, "Engr Approved By");
                case kDocumnetProperty.EngrDateApproved:
                    return GetProperty(doc, trackSet, "Engr Date Approved");
                case kDocumnetProperty.ExternalPropertyRevisionId:
                    return GetProperty(doc, trackSet, "External Property Revision Id");
                case kDocumnetProperty.Language:
                    return GetProperty(doc, trackSet, "Language");
                case kDocumnetProperty.Manufacturer:
                    return GetProperty(doc, trackSet, "Manufacturer");
                case kDocumnetProperty.Material:
                    return GetProperty(doc, trackSet, "Material");
                case kDocumnetProperty.MfgApprovedBy:
                    return GetProperty(doc, trackSet, "Mfg Approved By");
                case kDocumnetProperty.MfgDateApproved:
                    return GetProperty(doc, trackSet, "Mfg Date Approved");
                case kDocumnetProperty.ParameterizedTemplate:
                    return GetProperty(doc, trackSet, "Parameterized Template");
                case kDocumnetProperty.PartIcon:
                    return GetProperty(doc, trackSet, "Part Icon");
                case kDocumnetProperty.PartNumber:
                    return GetProperty(doc, trackSet, "Part Number");
                case kDocumnetProperty.PartPropertyRevisionId:
                    return GetProperty(doc, trackSet, "Part Property Revision Id");
                case kDocumnetProperty.Project:
                    return GetProperty(doc, trackSet, "Project");
                case kDocumnetProperty.ProxyRefreshDate:
                    return GetProperty(doc, trackSet, "Proxy Refresh Date");
                case kDocumnetProperty.SizeDesignation:
                    return GetProperty(doc, trackSet, "Size Designation");
                case kDocumnetProperty.Standard:
                    return GetProperty(doc, trackSet, "Standard");
                case kDocumnetProperty.StandardRevision:
                    return GetProperty(doc, trackSet, "Standard Revision");
                case kDocumnetProperty.StandardsOrganization:
                    return GetProperty(doc, trackSet, "Standards Organization");
                case kDocumnetProperty.StockNumber:
                    return GetProperty(doc, trackSet, "Stock Number");
                case kDocumnetProperty.TemplateRow:
                    return GetProperty(doc, trackSet, "Template Row");
                case kDocumnetProperty.UserStatus:
                    return GetProperty(doc, trackSet, "User Status");
                case kDocumnetProperty.Vendor:
                    return GetProperty(doc, trackSet, "Vendor");
                case kDocumnetProperty.WeldMaterial:
                    return GetProperty(doc, trackSet, "WeldMaterial");
                case kDocumnetProperty.Category:
                    return GetProperty(doc, docSet, "Category");
                case kDocumnetProperty.Company:
                    return GetProperty(doc, docSet, "Company");
                case kDocumnetProperty.Manager:
                    return GetProperty(doc, docSet, "Manager");
                case kDocumnetProperty.Author:
                    return GetProperty(doc, sumSet, "Author");
                case kDocumnetProperty.Comments:
                    return GetProperty(doc, sumSet, "Comments");
                case kDocumnetProperty.Keywords:
                    return GetProperty(doc, sumSet, "Keywords");
                case kDocumnetProperty.LastSavedBy:
                    return GetProperty(doc, sumSet, "Last Saved By");
                case kDocumnetProperty.Thumbnail:
                    return GetProperty(doc, sumSet, "Thumbnail");
                case kDocumnetProperty.RevisionNumber:
                    return GetProperty(doc, sumSet, "Revision Number");
                case kDocumnetProperty.Subject:
                    return GetProperty(doc, sumSet, "Subject");
                case kDocumnetProperty.Title:
                    return GetProperty(doc, sumSet, "Title");
                default:
                    return null;

            }
        }

        /// <summary>
        /// get a user defined property if it exists
        /// </summary>
        /// <param name="doc">Inventor.Document object</param>
        /// <param name="name">name of the property</param>
        /// <returns></returns>
        public static Property GetCustomProperty(this Document doc, string name)
        {
            return GetProperty(doc, "Inventor User Defined Properties", name);
        }

        /// <summary>
        /// assign a vlue to an inventor property
        /// </summary>
        /// <param name="doc">inventor.document object</param>
        /// <param name="property">kDocumnetProperty enumeration</param>
        /// <param name="value">the value to assign to property</param>
        public static void SetProperty(this Document doc, kDocumnetProperty property, object value)
        {
            Property emptyProp = GetProperty(doc, property);
            Type tp = emptyProp.Value.GetType();
            if (tp.Equals(value.GetType()))
            {
                emptyProp.Value = value;
            }
            else
            {
                throw new TypeAccessException($"The type of {value.ToString()} does not match {property.ToString()} type. " +
                    $"it should be of type {tp.ToString()}");
            }
        }

        /// <summary>
        /// assign a vlue to a user defined property
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="property"></param>
        /// <param name="value"></param>
        public static void SetCustomProperty(this Document doc, string property, object value)
        {
            // Get the custom property set. 
            PropertySet customPropSet = doc.PropertySets["Inventor User Defined Properties"];
            Property emptyProp;

            //see if custom property exists
            try
            {
                emptyProp = GetCustomProperty(doc, property);
            }

            //property does not exist so make it
            catch (Exception)
            {
                emptyProp = customPropSet.Add(value, property);
            }

            //assign value to property if types match
            Type tp = emptyProp.Value.GetType();
            if (tp.Equals(value.GetType()))
            {
                emptyProp.Value = value;
            }
            else
            {
                throw new TypeAccessException($"The type of {value.ToString()} does not match {property.ToString()} type. " +
                    $"it should be of type {tp.ToString()}");
            }
        }
    }
}
