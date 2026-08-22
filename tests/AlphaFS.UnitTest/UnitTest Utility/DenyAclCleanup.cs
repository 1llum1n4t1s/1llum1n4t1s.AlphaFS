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
using System.Security.AccessControl;
using System.Security.Principal;

namespace AlphaFS.UnitTest
{
   /// <summary>拒否 ACE や ReadOnly 属性が残ったテスト用ディレクトリを、確実に削除するための後始末。</summary>
   internal static class DenyAclCleanup
   {
      /// <summary>ディレクトリを再帰的に削除する。削除を妨げる拒否 ACE と ReadOnly 属性は道中で取り除く。</summary>
      /// <remarks>
      ///   拒否 ACE を張ったテストが finally での解除に失敗すると、%TEMP% に二度と消せない
      ///   AlphaFS.TempRoot.* が残り続ける (2026-07-26 のリリース作業で 21 個滞留した)。
      ///   拒否 FullControl はディレクトリの列挙そのものを塞ぐため、中を見てから消すことはできない。
      ///   「その階層を列挙する前に、その階層の拒否 ACE を外す」トップダウン走査だけが最後まで辿れる。
      ///   ディレクトリの所有者は DACL の内容に関わらず READ_CONTROL と WRITE_DAC を暗黙に持つので、
      ///   テストが自分で作ったディレクトリであれば拒否 ACE 下でも DACL は書き換えられる。
      /// </remarks>
      public static void ForceDelete(string directoryFullPath)
      {
         if (!System.IO.Directory.Exists(directoryFullPath))
         {
            return;
         }

         StripDeleteBlockers(directoryFullPath);

         foreach (var subDirectory in System.IO.Directory.EnumerateDirectories(directoryFullPath))
         {
            // 再解析ポイント (ジャンクション / シンボリックリンク) の先はテストの持ち物とは限らない。
            // 辿らずにリンク自体だけを消す。
            if (IsReparsePoint(subDirectory))
            {
               StripDeleteBlockers(subDirectory);
               System.IO.Directory.Delete(subDirectory);
            }

            else
            {
               ForceDelete(subDirectory);
            }
         }

         foreach (var file in System.IO.Directory.EnumerateFiles(directoryFullPath))
         {
            StripDeleteBlockers(file);
            System.IO.File.Delete(file);
         }

         System.IO.Directory.Delete(directoryFullPath);
      }


      private static bool IsReparsePoint(string fullPath)
      {
         try
         {
            return (System.IO.File.GetAttributes(fullPath) & System.IO.FileAttributes.ReparsePoint) != 0;
         }
         catch (Exception)
         {
            return false;
         }
      }


      /// <summary>削除を妨げる ReadOnly 属性と拒否 ACE を取り除く。取り除けなくても、後続の削除に判断を委ねる。</summary>
      private static void StripDeleteBlockers(string fullPath)
      {
         try
         {
            var attributes = System.IO.File.GetAttributes(fullPath);

            if ((attributes & System.IO.FileAttributes.ReadOnly) != 0)
            {
               System.IO.File.SetAttributes(fullPath, attributes & ~System.IO.FileAttributes.ReadOnly);
            }
         }
         catch (Exception)
         {
         }

         try
         {
            RemoveDenyRules(fullPath);
         }
         catch (Exception)
         {
         }
      }


      private static void RemoveDenyRules(string fullPath)
      {
         var isDirectory = (System.IO.File.GetAttributes(fullPath) & System.IO.FileAttributes.Directory) != 0;

         if (isDirectory)
         {
            var directoryInfo = new System.IO.DirectoryInfo(fullPath);

            // SACL は SE_SECURITY_NAME 特権が要るうえ、削除には無関係。DACL だけを読み書きする。
            var security = System.IO.FileSystemAclExtensions.GetAccessControl(directoryInfo, AccessControlSections.Access);

            if (TryRemoveDenyRules(security))
            {
               System.IO.FileSystemAclExtensions.SetAccessControl(directoryInfo, security);
            }
         }

         else
         {
            var fileInfo = new System.IO.FileInfo(fullPath);

            var security = System.IO.FileSystemAclExtensions.GetAccessControl(fileInfo, AccessControlSections.Access);

            if (TryRemoveDenyRules(security))
            {
               System.IO.FileSystemAclExtensions.SetAccessControl(fileInfo, security);
            }
         }
      }


      /// <summary>明示的に付与された拒否 ACE をすべて取り除く。継承した ACE は親を消せば消えるので触らない。</summary>
      /// <returns>1 件でも取り除いたら <c>true</c>。1 件も無ければ <c>false</c> (無駄な DACL 書き込みを避ける)。</returns>
      private static bool TryRemoveDenyRules(FileSystemSecurity security)
      {
         var removedAny = false;

         foreach (FileSystemAccessRule rule in security.GetAccessRules(true, false, typeof(SecurityIdentifier)))
         {
            if (rule.AccessControlType != AccessControlType.Deny)
            {
               continue;
            }

            security.RemoveAccessRuleSpecific(rule);

            removedAny = true;
         }

         return removedAny;
      }
   }
}
