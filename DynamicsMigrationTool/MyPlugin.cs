using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace DynamicsMigrationTool
{
    // Do not forget to update version number and author (company attribute) in AssemblyInfo.cs class
    // To generate Base64 string for Images below, you can use https://www.base64-image.de/
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Dynamics Migration Tool"),
        ExportMetadata("Description", "A tool to assist with the Migration of Data into Dynamics using an MSSQL Staging Database"),
        // Please specify the base64 content of a 32x32 pixels image
        ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAABGhJREFUWEftlntMlXUYxz/vey4gyM0bchFOYIKomKKQ3Bos54ABolYzWQ7NVpoCM1mCaRmTODAhD8sRqQtMZhm6vE28lltNpmbW3KJo0lp2AfFyQDhw3re9nBQOyOWAVH/0/v17n+fzfJ/v+31/gizLMv/iI/wP8N9RoNMJMheu/kz1+UtsWLMIe8FijnaTxN4jNaQtevKRu0XI2VIgf/XFl6ydH8zhH4xs267H3VW0alR9/mviVxbQcu0jtOq/qR4RihAdNV+e4uNN7a+NFBTpCZ0Z0Lu0LCF6xZCcvJCDOzM7lWoyduDmqIFh8giTpsXKTc33MLa0IGjsqL94gkkTXa0gZCS0/okIbc0Y9NkkxYagi07j1IEPiAqe0I8Wyl77JxRE71gZtSOy6RYiHUwPms6Vk+/3Krr7kzOsSs/CTuuIt8cY6up/I27xKo4YVjwUQJJM7Dv+DakJc/tdlpBVfFTelh6H2SwQFLuWn67XUvFeLst6vPj99d8JikjEQQ1au1E4O9qzcf06XkpNsBjVLKFRWbwjI5OZu4+SsnLOVhm402JknJMjYU/0Xq9VEDUZTejLz3H09CWufroBUD+glyUZrX88YusNgv0moBbMnP38dOeXsjp3L888HUpM+JQH5zcVHyIvPx97e0c6Oto4UF5MYkxILzV6JOH9nSkzCD22J7Hr42o2bs5D52xmjp8nATpfGscE83pGKg5a613faLiNz6xk7EUTo0epWLxwCSXvpHcN1OkOGZujOOG5l3GR7uIz3o3mcbMxbH24B5QRnALiEO41EOLvxozAORh25nUCtJnaEUUNapUNAMovq1mCk8fOcLddIHCKH6HTfPs0uaJlnqGSwh2lzNVpGT/aibQX0pgdGcHxmmvM9Z/ERJ237QrYkj+KCr5TI0mc4UWA+xg85ywgZXkKqm5FbF6BLQDtrXfIz3mNxtsyScuWMy8iHHutdcqOGIAZmU2FlWxd/zxKelss2juYRgygj34DfYa2CGw523SrhQOnLhI7L5jHvFyxFnjgerYrIIMsQMbmEqS6S8SvzCQmYipajRpRtP3PZBOAssHSimMU7tiF2XQLf/exVFSW4THWZeBR+zgxaICGm0YWrX6bm42tXPuuBlHJfXMzOPnRUVs18gDdO4QnrebCxcuoVCKSbMeVc/uYPtljSBADKqDcBQQra8m8u/sQmTlFhIXNpKbmMvv3bGfJgrDhAShRu+aNUm782UTGslgio0JJ31KGp9so1r2aymhNV33l7IK0PA6XZTE58kWWLn0WfUbc0AGUKfWlR8nems8sX2eyE4JpcnmcFWtWIghmsArP+30soaLAXK5rJGTy2KED+Ea9QmvrbUwaD1obfkS814DoIPDt6Sp03v1duYbU0+oliwe6JeQvfxjJeKuEqkMn8PbSUV+zZ7j3zn4pH2pChSe7+CD6YgMfFr1Jakr08Ee1NQcUiP2fncHBxZOkpwL/eYCujhLYnPCD5x0wBwZfamgn/wLnXNCVVrRC4wAAAABJRU5ErkJggg=="),
        // Please specify the base64 content of a 80x80 pixels image
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]
    public class MyPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MyPluginControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public MyPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>

        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}