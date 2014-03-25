using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K2Field.SmartObjects.Services.SharePoint.Search.Data
{
    public class StandardReturns
    {
        public static List<Property> GetStandardReturnProperties()
        {
            List<Property> StandardReturnProperties = new List<Property>();

            //StandardReturnProperties.AddRange(GetStandardInputProperties());

            Property responsestatus = new Property();
            responsestatus.Name = "responsestatus";
            responsestatus.MetaData.DisplayName = "Response Status";
            responsestatus.SoType = SoType.Text;
            StandardReturnProperties.Add(responsestatus);

            Property responsestatusdescription = new Property();
            responsestatusdescription.Name = "responsestatusdescription";
            responsestatusdescription.MetaData.DisplayName = "Response Status Description";
            responsestatusdescription.SoType = SoType.Memo;
            StandardReturnProperties.Add(responsestatusdescription);

            return StandardReturnProperties;

        }
    }
}
