using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AccessCodeLib.DeclarationReader
{
    [ComVisible(true)]
    [Guid("50065C31-0430-4AF3-9CBA-AC20B9357A9D")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ACLibDeclarationReader.TypeLibDeclarationReader")]
    public class TypeLibDeclarationReader : ITypeLibDeclarationReader
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(
            [MarshalAs(UnmanagedType.LPWStr)] string strTypeLibName,
            REGKIND regKind,
            out ITypeLib typeLib);

        private enum REGKIND
        {
            REGKIND_DEFAULT = 0,
            REGKIND_REGISTER = 1,
            REGKIND_NONE = 2
        }

        public string[] ReadTypeLib(string typeLibFilePath)
        {
            LoadTypeLibEx(typeLibFilePath, REGKIND.REGKIND_NONE, out ITypeLib typeLib);

            var declarations = new List<string>();
            int typeCount = typeLib.GetTypeInfoCount();

            for (int i = 0; i < typeCount; i++)
            {
                typeLib.GetTypeInfo(i, out ITypeInfo typeInfo);
                typeInfo.GetTypeAttr(out IntPtr pTypeAttr);
                var typeAttr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                typeInfo.GetDocumentation(-1, out string typeName, out _, out _, out _);
                declarations.Add(typeName);

                for (int j = 0; j < typeAttr.cFuncs; j++)
                {
                    typeInfo.GetFuncDesc(j, out IntPtr pFuncDesc);
                    var funcDesc = (System.Runtime.InteropServices.ComTypes.FUNCDESC)Marshal.PtrToStructure(pFuncDesc, typeof(System.Runtime.InteropServices.ComTypes.FUNCDESC));

                    typeInfo.GetDocumentation(funcDesc.memid, out string methodName, out _, out _, out _);
                    declarations.Add(methodName);

                    typeInfo.ReleaseFuncDesc(pFuncDesc);
                }

                for (int k = 0; k < typeAttr.cVars; k++)
                {
                    typeInfo.GetVarDesc(k, out IntPtr pVarDesc);
                    var varDesc = (System.Runtime.InteropServices.ComTypes.VARDESC)Marshal.PtrToStructure(pVarDesc, typeof(System.Runtime.InteropServices.ComTypes.VARDESC));

                    typeInfo.GetDocumentation(varDesc.memid, out string propName, out _, out _, out _);
                    declarations.Add(propName);

                    typeInfo.ReleaseVarDesc(pVarDesc);
                }

                typeInfo.ReleaseTypeAttr(pTypeAttr);
            }

            return declarations.Distinct().ToArray();
        }
    }

    [ComVisible(true)] // Interface for COM visibility
    [Guid("2031C42E-4C73-4722-857E-8043D6B1CC4C")]
    public interface ITypeLibDeclarationReader
    {
        string[] ReadTypeLib(string TypeLibFilePath);
    }
}
