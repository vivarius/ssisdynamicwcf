using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ServiceModel.Description;
using System.Text;

namespace SSISWCFTask100.WCFProxy
{
    public class DynamicProxyException : ApplicationException
    {
        public DynamicProxyException(string message)
            : base(message)
        {
        }

        public DynamicProxyException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public IEnumerable<MetadataConversionError> MetadataImportErrors { get; internal set; }

        public IEnumerable<MetadataConversionError> CodeGenerationErrors { get; internal set; }

        public IEnumerable<CompilerError> CompilationErrors { get; internal set; }

        public override string ToString()
        {
            var builder = new StringBuilder();
            builder.AppendLine(base.ToString());

            if (MetadataImportErrors != null)
            {
                builder.AppendLine("Metadata Import Errors:");
                builder.AppendLine(DynamicProxyFactory.ToString(MetadataImportErrors));
            }

            if (CodeGenerationErrors != null)
            {
                builder.AppendLine("Code Generation Errors:");
                builder.AppendLine(DynamicProxyFactory.ToString(CodeGenerationErrors));
            }

            if (CompilationErrors != null)
            {
                builder.AppendLine("Compilation Errors:");
                builder.AppendLine(DynamicProxyFactory.ToString(CompilationErrors));
            }

            return builder.ToString();
        }
    }
}