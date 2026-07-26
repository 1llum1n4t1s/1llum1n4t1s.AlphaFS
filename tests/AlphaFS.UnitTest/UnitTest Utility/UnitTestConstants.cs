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

using System;
using System.Globalization;

namespace AlphaFS.UnitTest
{
   /// <summary>Containts static variables, used by unit tests.</summary>
   public static partial class UnitTestConstants
   {
      /// <summary>The Computer system drive. For example "C:".</summary>
      public static readonly string SysDrive = Environment.GetEnvironmentVariable("SystemDrive");


      public const int OneMebibyte = 1048576;


      public static readonly string[] StreamArrayContent =
      {
         "(1) Nikolai Tesla: \"Today's scientists have substituted mathematics for experiments, and they wander off through equation after equation, and eventually build a structure which has no relation to reality.\"",
         "(2) The quick brown fox jumps over the lazy dog.",
         "(3) Computer: [" + Environment.MachineName + "]" + "\tHello there, " + Environment.UserName
      };


      /// <summary>
      ///   AlphaFS が System.IO と意図的に異なる結果を返す入力パスかどうかを判定する。
      /// </summary>
      /// <remarks>
      ///   AlphaFS は拡張長プレフィックス (<c>\\?\</c>) を解決・除去して通常形式のパスを返すのが仕様であり、
      ///   プレフィックスを保持する System.IO とは結果が一致しない。また UNC ルート (<c>\\server\share</c>) を
      ///   ルートではなく 1 つの構成要素として扱うため、GetFileName 系の結果も一致しない。
      ///   これらは不具合ではなくライブラリの目的そのものなので、System.IO との等価表明の対象から外す。
      /// </remarks>
      public static bool DivergesFromSystemIo(string path)
      {
         if (Alphaleonis.Utils.IsNullOrWhiteSpace(path))
         {
            return false;
         }

         // 拡張長プレフィックス (\\?\ および \\?\UNC\) を含むパス。
         if (path.StartsWith(Alphaleonis.Win32.Filesystem.Path.LongPathPrefix, StringComparison.Ordinal))
         {
            return true;
         }

         // UNC ルート (\\server\share) そのもの。末尾に構成要素が無いものだけが対象。
         if (path.StartsWith(Alphaleonis.Win32.Filesystem.Path.UncPrefix, StringComparison.Ordinal))
         {
            var parts = path.Substring(Alphaleonis.Win32.Filesystem.Path.UncPrefix.Length)
               .Split(new[] { Alphaleonis.Win32.Filesystem.Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);

            return parts.Length <= 2;
         }

         return false;
      }


      public static readonly string[] InputPaths =
      {
         @".",
         @".zip",

         SysDrive + @"\\test.txt",
         SysDrive + @"\/test.txt",

         System.IO.Path.DirectorySeparatorChar.ToString(),
         System.IO.Path.DirectorySeparatorChar + @"Program Files\Microsoft Office",

         Alphaleonis.Win32.Filesystem.Path.GlobalRootDevicePrefix + @"harddisk0\partition1\",
         Alphaleonis.Win32.Filesystem.Path.VolumePrefix + @"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}\Program Files\notepad.exe",

         "dir1/dir2/dir3/",

         @"Program Files\Microsoft Office",
         SysDrive[0].ToString(CultureInfo.InvariantCulture),
         SysDrive,
         SysDrive + @"\",
         SysDrive + @"\a",
         SysDrive + @"\a\",
         SysDrive + @"\a\b",
         SysDrive + @"\a\b\",
         SysDrive + @"\a\b\c",
         SysDrive + @"\a\b\c\",
         SysDrive + @"\a\b\c\z",
         SysDrive + @"\a\b\c\z.",
         SysDrive + @"\a\b\c\z.t",
         SysDrive + @"\a\b\c\z.tx",
         SysDrive + @"\a\b\c\z.txt",

         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + @"Program Files\Microsoft Office",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive[0].ToString(CultureInfo.InvariantCulture),
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive,
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\z",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\z.",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\z.t",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\z.tx",
         Alphaleonis.Win32.Filesystem.Path.LongPathPrefix + SysDrive + @"\a\b\c\z.txt",

         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\f",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\fi",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\fil",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\file",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\file.",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\file.e",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\file.ex",
         Alphaleonis.Win32.Filesystem.Path.UncPrefix + Environment.MachineName + @"\Share\d1\d2\file.ext",

         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\f",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\fi",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\fil",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\file",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\file.",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\file.e",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\file.ex",
         Alphaleonis.Win32.Filesystem.Path.LongPathUncPrefix + Environment.MachineName + @"\Share\d1\d2\file.ext"
      };
   }
}
