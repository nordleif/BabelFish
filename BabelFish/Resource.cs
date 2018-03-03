using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BabelFish
{
    public class Resource
    {
        public Resource()
        {

        }
        
        public string DestinationText { get; set; }
        
        public string ResourceName { get; set; }

        public string SourceFileName { get; set; }

        public string SourceText { get; set; }
    }
}
