using System;
using System.Linq;
using System.IO;
using System.Reflection;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.uno;
using System.Collections.Generic;

namespace UnoSharp
{
    public static class OfficeServiceManager
    {
        public static XComponentContext bootstrap { get; }
        public static XMultiServiceFactory factory { get; }
        public static XComponentLoader loader { get; }

        static OfficeServiceManager()
        {
            bootstrap = uno.util.Bootstrap.bootstrap();
            factory = (XMultiServiceFactory)bootstrap.getServiceManager();
            loader = (XComponentLoader)factory.createInstance("com.sun.star.frame.Desktop");
        }
    }
}
