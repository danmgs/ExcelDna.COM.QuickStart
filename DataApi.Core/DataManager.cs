using System;

namespace DataApi.Core
{
    public class DataManager
    {
        public string GetHello(string suffix)
        {
            return string.Format("Hello World {0}", suffix);
        }
    }
}
