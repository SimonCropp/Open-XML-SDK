// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Xml;

namespace DocumentFormat.OpenXml
{
    internal static class FrameworkExtensions
    {
        public static bool IsNullOrEmpty([NotNullWhen(false)] this string? str)
            => string.IsNullOrEmpty(str);

        /// <summary>
        /// Returns true if the string is a valid NCName, false otherwise.
        /// Unlike <see cref="XmlConvert.VerifyNCName"/>, this does not throw on invalid input.
        /// </summary>
        public static bool IsNCName([NotNullWhen(true)] this string? name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            try
            {
                XmlConvert.VerifyNCName(name);
                return true;
            }
            catch (XmlException)
            {
                return false;
            }
        }

        public static MemoryStream CopyToMemoryStream(this Stream stream)
        {
            if (stream.Length > int.MaxValue)
            {
                throw new OpenXmlPackageException(ExceptionMessages.DocumentTooBig);
            }

            var memoryStream = new MemoryStream(Convert.ToInt32(stream.Length));
            stream.CopyTo(memoryStream);

            return memoryStream;
        }

        public static MemoryStream CopyToMemoryStream(this OpenXmlPart part)
        {
            using (var stream = part.GetStream())
            {
                return stream.CopyToMemoryStream();
            }
        }
    }
}
