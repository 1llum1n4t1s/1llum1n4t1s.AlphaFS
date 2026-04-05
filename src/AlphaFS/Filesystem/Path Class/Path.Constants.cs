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
using System.Diagnostics.CodeAnalysis;
using System.Globalization;

namespace Alphaleonis.Win32.Filesystem
{
   public static partial class Path
   {
      /// <summary>[AlphaFS] SearchPatternからトリムする文字。</summary>
      internal static readonly char[] TrimEndChars = {(char) 0x9, (char) 0xA, (char) 0xB, (char) 0xC, (char) 0xD, (char) 0x20, (char) 0x85, (char) 0xA0};

      /// <summary>AltDirectorySeparatorChar = '/' 階層的なファイルシステム構成を反映するパス文字列内のディレクトリ階層を区切るための、プラットフォーム固有の代替文字を提供します。</summary>
      public static readonly char AltDirectorySeparatorChar = System.IO.Path.AltDirectorySeparatorChar;

      /// <summary>[AlphaFS] AltDirectorySeparatorChar = "/" 階層的なファイルシステム構成を反映するパス文字列内のディレクトリ階層を区切るための、プラットフォーム固有の代替文字列を提供します。</summary>
      public static readonly string AltDirectorySeparator = AltDirectorySeparatorChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>DirectorySeparatorChar = '\' 階層的なファイルシステム構成を反映するパス文字列内のディレクトリ階層を区切るための、プラットフォーム固有の文字を提供します。</summary>
      public static readonly char DirectorySeparatorChar = System.IO.Path.DirectorySeparatorChar;

      /// <summary>[AlphaFS] DirectorySeparator = "\" 階層的なファイルシステム構成を反映するパス文字列内のディレクトリ階層を区切るための、プラットフォーム固有の文字列を提供します。</summary>
      public static readonly string DirectorySeparator = DirectorySeparatorChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>[AlphaFS] NetworkDriveSeparator = '$' プラットフォーム固有のネットワークドライブ区切り文字を提供します。</summary>
      public const char NetworkDriveSeparatorChar = '$';

      /// <summary>[AlphaFS] NetworkDriveSeparator = "$" プラットフォーム固有のネットワークドライブ区切り文字列を提供します。</summary>
      public static readonly string NetworkDriveSeparator = NetworkDriveSeparatorChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>PathSeparator = ';' 環境変数内のパス文字列を区切るためのプラットフォーム固有の区切り文字。</summary>
      public static readonly char PathSeparator = System.IO.Path.PathSeparator;


      /// <summary>VolumeSeparatorChar = ':' プラットフォーム固有のボリューム区切り文字を提供します。</summary>
      public static readonly char VolumeSeparatorChar = System.IO.Path.VolumeSeparatorChar;

      /// <summary>[AlphaFS] VolumeSeparator = ":" プラットフォーム固有のボリューム区切り文字列を提供します。</summary>
      public static readonly string VolumeSeparator = VolumeSeparatorChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>[AlphaFS] StreamSeparator = ':' プラットフォーム固有のストリーム名文字を提供します。</summary>
      public const char StreamSeparatorChar = ':';

      /// <summary>[AlphaFS] StreamSeparator = ':' プラットフォーム固有のストリーム名文字列を提供します。</summary>
      public static readonly string StreamSeparator = StreamSeparatorChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>[AlphaFS] StreamDataLabel = ':$DATA' プラットフォーム固有のストリーム :$DATA ラベルを提供します。</summary>
      public static readonly string StreamDataLabel = StreamSeparator + "$DATA";

      /// <summary>[AlphaFS] StringTerminatorChar = '\0' 文字列終端サフィックス。</summary>
      public const char StringTerminatorChar = '\0';


      /// <summary>[AlphaFS] CurrentDirectoryPrefix = '.' カレントディレクトリ文字を提供します。</summary>
      public const char CurrentDirectoryPrefixChar = '.';

      /// <summary>[AlphaFS] CurrentDirectoryPrefix = "." カレントディレクトリ文字列を提供します。</summary>
      public static readonly string CurrentDirectoryPrefix = CurrentDirectoryPrefixChar.ToString(CultureInfo.InvariantCulture);

      /// <summary>[AlphaFS] ExtensionSeparatorChar = '.' 拡張子区切り文字を提供します。</summary>
      public const char ExtensionSeparatorChar = '.';

      /// <summary>[AlphaFS] ParentDirectoryPrefix = ".." 親ディレクトリ文字列を提供します。</summary>
      public const string ParentDirectoryPrefix = "..";

      /// <summary>[AlphaFS] WildcardStarMatchAll = '*' 全項目一致ワイルドカード文字を提供します。</summary>
      public const char WildcardStarMatchAllChar = '*';

      /// <summary>[AlphaFS] WildcardStarMatchAll = "*" 全項目一致ワイルドカード文字列を提供します。</summary>
      public static readonly string WildcardStarMatchAll = WildcardStarMatchAllChar.ToString(CultureInfo.InvariantCulture);

      /// <summary>[AlphaFS] WildcardQuestion = '?' 単一文字置換ワイルドカード文字を提供します。</summary>
      public const char WildcardQuestionChar = '?';

      /// <summary>[AlphaFS] WildcardQuestion = "?" 単一文字置換ワイルドカード文字列を提供します。</summary>
      public static readonly string WildcardQuestion = WildcardQuestionChar.ToString(CultureInfo.InvariantCulture);


      /// <summary>[AlphaFS] Win32 ファイル名前空間。パス文字列の "\\?\" プレフィックスは、Windows APIに対してすべての文字列解析を無効にし、それに続く文字列をそのままファイルシステムに送信するよう指示します。</summary>
      public static readonly string LongPathPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{0}{1}{0}", DirectorySeparatorChar, WildcardQuestion);

      /// <summary>[AlphaFS] Win32 デバイス名前空間。"\\.\" プレフィックスは、APIがこの種のアクセスをサポートしている場合に、ファイルシステムを経由せずに物理ディスクやボリュームにアクセスする方法です。</summary>
      public static readonly string LogicalDrivePrefix = string.Format(CultureInfo.InvariantCulture, "{0}{0}.{0}", DirectorySeparatorChar);

      /// <summary>[AlphaFS] PhysicalDrivePrefix = "\\.\PhysicalDrive" 標準的な物理ドライブプレフィックスを提供します。</summary>
      public static readonly string PhysicalDrivePrefix = string.Format(CultureInfo.InvariantCulture, "{0}PhysicalDrive", LogicalDrivePrefix);


      /// <summary>[AlphaFS] GlobalRootPrefix = "\\?\GlobalRoot\" 標準的なWindowsボリュームプレフィックスを提供します。</summary>
      public static readonly string GlobalRootPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", LongPathPrefix, "GlobalRoot", DirectorySeparatorChar);

      /// <summary>[AlphaFS] GlobalRootDevicePrefix = "\\?\GlobalRoot\Device\" 標準的なWindowsボリュームプレフィックスを提供します。</summary>
      public static readonly string GlobalRootDevicePrefix = string.Format(CultureInfo.InvariantCulture, "{0}{2}{1}{3}{1}", LongPathPrefix, DirectorySeparatorChar, "GlobalRoot", "Device");

      /// <summary>[AlphaFS] NonInterpretedPathPrefix = "\??\" 非解釈パスプレフィックスを提供します。</summary>
      public static readonly string NonInterpretedPathPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{1}{0}", DirectorySeparatorChar, WildcardQuestion);

      /// <summary>[AlphaFS] VolumePrefix = "\\?\Volume" 標準的なWindowsボリュームプレフィックスを提供します。</summary>
      public static readonly string VolumePrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}", LongPathPrefix, "Volume");

      /// <summary>[AlphaFS] DevicePrefix = "\Device\" 標準的なWindowsデバイスプレフィックスを提供します。</summary>
      public static readonly string DevicePrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{0}", DirectorySeparatorChar, "Device");

      /// <summary>[AlphaFS] DosDeviceLanmanPrefix = "\Device\LanmanRedirector\" ネットワーク共有へのMS-Dos Lanmanリダイレクタ パス UNCプレフィックスを提供します。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Lanman")]
      [Obsolete("Unused")]
      public static readonly string DosDeviceLanmanPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", DevicePrefix, "LanmanRedirector", DirectorySeparatorChar);

      /// <summary>[AlphaFS] DosDeviceMupPrefix = "\Device\Mup\" ネットワーク共有へのMS-Dos Mupリダイレクタ パス UNCプレフィックスを提供します。</summary>
      [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Mup")]
      [Obsolete("Unused")]
      public static readonly string DosDeviceMupPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", DevicePrefix, "Mup", DirectorySeparatorChar);


      /// <summary>[AlphaFS] UncPrefix = "\\" 標準的なWindows パス UNCプレフィックスを提供します。</summary>
      public static readonly string UncPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{0}", DirectorySeparatorChar);

      /// <summary>[AlphaFS] DosDeviceUncPrefix = "\??\UNC\" ネットワーク共有へのSUBST.EXE パス UNCプレフィックスを提供します。</summary>
      public static readonly string DosDeviceUncPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", NonInterpretedPathPrefix, "UNC", DirectorySeparatorChar);

      /// <summary>[AlphaFS] LongPathUncPrefix = "\\?\UNC\" 標準的なWindows長パス UNCプレフィックスを提供します。</summary>
      public static readonly string LongPathUncPrefix = string.Format(CultureInfo.InvariantCulture, "{0}{1}{2}", LongPathPrefix, "UNC", DirectorySeparatorChar);
   }
}
