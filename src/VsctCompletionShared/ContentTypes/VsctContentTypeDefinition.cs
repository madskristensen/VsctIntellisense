﻿using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Utilities;

namespace VsctCompletion.ContentTypes
{
    public class VsctContentTypeDefinition
    {
        public const string VsctContentType = "VSCT";

        [Export(typeof(ContentTypeDefinition))]
        [Name(VsctContentType)]
        [BaseDefinition("XML")]
        public ContentTypeDefinition IVsctContentType { get; set; }

        [Export(typeof(FileExtensionToContentTypeDefinition))]
        [ContentType(VsctContentType)]
        [FileExtension(".vsct")]
        public FileExtensionToContentTypeDefinition VsctFileExtension { get; set; }
    }
}
