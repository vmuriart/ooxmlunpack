namespace OoXmlUnpack
{
    using System.Xml.Linq;

    public static class Extensions
    {
        public static void ChangeOrAddAttribute(this XElement element, string attributeToChangeOrAdd, string valueToSetForAttribute)
        {
            var attribute = element.Attribute(attributeToChangeOrAdd);
            if (attribute != null)
            {
                attribute.Value = valueToSetForAttribute;
            }
            else
            {
                element.Add(new XAttribute(attributeToChangeOrAdd, valueToSetForAttribute));
            }
        }
    }
}