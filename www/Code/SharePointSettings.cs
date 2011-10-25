using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;

namespace SharePointConnector
{
    /// <summary>
    /// Handles entries within the web.config which relate to adding of new List Types
    /// </summary>
    /// 

    public class SharePointSettings : ConfigurationSection
    {
        [ConfigurationProperty("SharePointLists")]
        public SharePointLists SharepointLists
        {
            get
            {
                return this["SharePointLists"] as SharePointLists;
            }
        }
    }

    
    public class SharePointLists : ConfigurationElementCollection
    {
        public List this[int index]
        {
            get
            {
                return base.BaseGet(index) as List;
            }
        }

        protected override ConfigurationElement CreateNewElement()
        {
            return new List();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            if (element == null) throw new ArgumentException("element");
                return ((List)element).Name;
        }
    }


    public class List : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get 
            { 
                return (string)this["name"]; 
            }
        }

        [ConfigurationProperty("templateId", IsRequired = true)]
        public string TemplateId
        {
            get
            {
                return (string)this["templateId"];
            }
        }
    }


    
}
