/*  Copyright (C) 2008-2018 Peter Palotas, Jeffrey Jangli, Alexandr Normuradov
 *  
 *  Permission is hereby granted, free of charge, to any person obtaining a copy 
 *  of this software and associated documentation files (the "Software"), to deal 
 *  in the Software without restriction, including without limitation the rights 
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
 *  copies of the Software, and to permit persons to whom the Software is 
 *  furnished to do so, subject to the following conditions:
 *  
 *  The above copyright notice and this permission notice shall be included in 
 *  all copies or substantial portions of the Software.
 *  
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
 *  THE SOFTWARE. 
 */

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace AlphaFS.UnitTest
{
   public partial class MoveTest
   {
      // Pattern: <class>_<function>_<scenario>_<expected result>


      [TestMethod]
      public void AlphaFS_Directory_Move_Overwrite_DestinationDirectoryAlreadyExists_LocalAndNetwork_Success()
      {
         AlphaFS_Directory_Move_Overwrite_DestinationDirectoryAlreadyExists(false);
         AlphaFS_Directory_Move_Overwrite_DestinationDirectoryAlreadyExists(true);
      }


      private void AlphaFS_Directory_Move_Overwrite_DestinationDirectoryAlreadyExists(bool isNetwork)
      {
         using (var tempRoot = new TemporaryDirectory(isNetwork))
         {
            var folderSrc = tempRoot.CreateTree();
            var folderDst = tempRoot.RandomDirectoryFullPath;

            Console.WriteLine("Src Directory Path: [{0}]", folderSrc.FullName);
            Console.WriteLine("Dst Directory Path: [{0}]", folderDst);
            

            Alphaleonis.Win32.Filesystem.Directory.Copy(folderSrc.FullName, folderDst);


            var dirEnumOptions = Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.FilesAndFolders | Alphaleonis.Win32.Filesystem.DirectoryEnumerationOptions.Recursive;

            var props = Alphaleonis.Win32.Filesystem.Directory.GetProperties(folderSrc.FullName, dirEnumOptions);
            var sourceTotal = props["Total"];
            var sourceTotalFiles = props["File"];
            var sourceTotalSize = props["Size"];

            Console.WriteLine("\n\tTotal size: [{0}] - Total Folders: [{1}] - Files: [{2}]", Alphaleonis.Utils.UnitSizeToText(sourceTotalSize), sourceTotal - sourceTotalFiles, sourceTotalFiles);


            // Overwrite using MoveOptions.ReplaceExisting

            Alphaleonis.Win32.Filesystem.Directory.Move(folderSrc.FullName, folderDst, Alphaleonis.Win32.Filesystem.MoveOptions.ReplaceExisting);


            props = Alphaleonis.Win32.Filesystem.Directory.GetProperties(folderDst, dirEnumOptions);
            Assert.AreEqual(sourceTotal, props["Total"], "The number of total file system objects does not match, but is expected to.");
            Assert.AreEqual(sourceTotalFiles, props["File"], "The number of total files does not match, but is expected to.");
            Assert.AreEqual(sourceTotalSize, props["Size"], "The total file size does not match, but is expected to.");
         }


         Console.WriteLine();
      }


      [TestMethod]
      public void AlphaFS_Directory_Move_Overwrite_SourceMoveFails_PreservesDestination_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var folderSrc = tempRoot.CreateTree();
         var folderDst = folderSrc.CreateSubdirectory("destination");
         var destinationMarker = System.IO.Path.Combine(folderDst.FullName, "destination-marker.txt");
         System.IO.File.WriteAllText(destinationMarker, "preserve me");

         UnitTestAssert.ThrowsException<System.IO.IOException>(() =>
            Alphaleonis.Win32.Filesystem.Directory.Move(folderSrc.FullName, folderDst.FullName, Alphaleonis.Win32.Filesystem.MoveOptions.ReplaceExisting));

         Assert.IsTrue(System.IO.Directory.Exists(folderSrc.FullName));
         Assert.IsTrue(System.IO.Directory.Exists(folderDst.FullName));
         Assert.AreEqual("preserve me", System.IO.File.ReadAllText(destinationMarker));
      }


      [TestMethod]
      public void AlphaFS_Directory_Move_Overwrite_EmulatedCopyFails_RestoresDestinationAndSource_Local_Success()
      {
         using var tempRoot = new TemporaryDirectory();
         var source = System.IO.Path.Combine(tempRoot.Directory.FullName, "Source");
         var destination = System.IO.Path.Combine(tempRoot.Directory.FullName, "Destination");
         var sourceMarker = System.IO.Path.Combine(source, "source-marker.txt");
         var destinationMarker = System.IO.Path.Combine(destination, "destination-marker.txt");

         System.IO.Directory.CreateDirectory(source);
         System.IO.Directory.CreateDirectory(destination);
         System.IO.File.WriteAllText(sourceMarker, "source");
         System.IO.File.WriteAllText(destinationMarker, "destination");

         var sourceLongPath = @"\\?\" + source;
         var destinationLongPath = @"\\?\" + destination;

         Assert.ThrowsExactly<InvalidOperationException>(() =>
            Alphaleonis.Win32.Filesystem.Directory.CopyMoveCore(new Alphaleonis.Win32.Filesystem.CopyMoveArguments
            {
               SourcePath = sourceLongPath,
               SourcePathLp = sourceLongPath,
               DestinationPath = destinationLongPath,
               DestinationPathLp = destinationLongPath,
               PathFormat = Alphaleonis.Win32.Filesystem.PathFormat.LongFullPath,
               PathsChecked = true,
               IsCopy = true,
               EmulateMove = true,
               MoveOptions = Alphaleonis.Win32.Filesystem.MoveOptions.ReplaceExisting,
               DirectoryEnumerationFilters = new Alphaleonis.Win32.Filesystem.DirectoryEnumerationFilters
               {
                  InclusionFilter = _ => throw new InvalidOperationException("copy failure")
               }
            }));

         Assert.AreEqual("source", System.IO.File.ReadAllText(sourceMarker));
         Assert.AreEqual("destination", System.IO.File.ReadAllText(destinationMarker));
         Assert.AreEqual(0, System.IO.Directory.GetFileSystemEntries(tempRoot.Directory.FullName, "Destination.alphafs-*").Length);
      }
   }
}
