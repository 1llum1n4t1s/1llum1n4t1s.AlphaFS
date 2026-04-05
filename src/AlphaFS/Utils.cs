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
using System.ComponentModel;
using System.Globalization;
using System.Text;

namespace Alphaleonis
{
   internal static class Utils
   {
      private static readonly string[] SizeFormats = {"B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"};

      // Source: https://stackoverflow.com/questions/6275980/string-replace-ignoring-case/45756981#45756981


      /// <summary>大文字と小文字の違いを無視して、現在のインスタンス内の指定された文字列のすべての出現箇所を別の指定された文字列で置き換えた新しい文字列を返します。</summary>
      internal static string ReplaceIgnoreCase(this string str, string oldValue, string newValue)
      {
         ArgumentNullException.ThrowIfNull(str);

         if (string.IsNullOrWhiteSpace(str))
         {
            return str;
         }

         ArgumentNullException.ThrowIfNull(oldValue);

         if (string.IsNullOrWhiteSpace(oldValue))
         {
            throw new ArgumentException("String cannot be of zero length.");
         }


         // 処理済み文字列を格納するための StringBuilder を準備する。
         var resultStringBuilder = new StringBuilder(str.Length);

         // 置換を分析: 置換か削除か。
         var isReplacementNullOrWhiteSpace = string.IsNullOrWhiteSpace(newValue);


         int foundAt;
         const int valueNotFound = -1;
         var startSearchFromIndex = 0;

         while ((foundAt = str.IndexOf(oldValue, startSearchFromIndex, StringComparison.OrdinalIgnoreCase)) != valueNotFound)
         {
            // 検出された置換箇所までのすべての文字を追加する。
            var charsUntilReplacment = foundAt - startSearchFromIndex;

            var isNothingToAppend = charsUntilReplacment == 0;

            if (!isNothingToAppend)
            {
               resultStringBuilder.Append(str, startSearchFromIndex, charsUntilReplacment);
            }

            if (!isReplacementNullOrWhiteSpace)
            {
               resultStringBuilder.Append(newValue);
            }


            // 次の検索の開始インデックスを準備する。
            // これは無限ループを防ぐために必要。そうしないとメソッドは常に文字列の先頭から検索を開始する。
            // 例: oldValue == "EXAMPLE"、newValue == "example"、
            // comparisonType == "大文字小文字無視" の場合、
            // "EXAMPLE" → "example" → "example" → "example" と無限ループになる。
            startSearchFromIndex = foundAt + oldValue.Length;

            if (startSearchFromIndex == str.Length)
            {
               // 入力文字列の終端: 次の検索のためのスペースがない。
               // 入力文字列は既に置換された値で終了している。
               // そのため、結果を持つ string builder は完成しており、追加のアクションは不要。
               return resultStringBuilder.ToString();
            }
         }


         // 最後の部分を結果に追加する。
         var charsUntilStringEnd = str.Length - startSearchFromIndex;

         resultStringBuilder.Append(str, startSearchFromIndex, charsUntilStringEnd);


         return resultStringBuilder.ToString();
      }


      /// <summary>列挙フィールド値の属性を取得します。</summary>
      /// <returns>列挙オプションに属する説明（文字列として）。</returns>
      /// <param name="enumValue"><see cref="Alphaleonis.Win32.Filesystem.DeviceGuid"/> 列挙型のいずれか。</param>
      public static string GetEnumDescription(Enum enumValue)
      {
         var enumValueString = enumValue.ToString();

         var fi = enumValue.GetType().GetField(enumValueString);

         if (fi is null)
         {
            return enumValueString;
         }

         var attributes = (DescriptionAttribute[]) fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

         return attributes.Length > 0 ? attributes[0].Description : enumValueString;
      }


      /// <summary>指定された文字列が null、空、または空白文字のみで構成されているかどうかを示します。</summary>
      /// <returns><paramref name="value"/> パラメーターが null または <see cref="string.Empty"/>、もしくは <paramref name="value"/> が空白文字のみで構成されている場合は <c>true</c>。</returns>
      /// <param name="value">テストする文字列。</param>
      public static bool IsNullOrWhiteSpace(string value)
      {
         return string.IsNullOrWhiteSpace(value);
      }


      /// <summary>型 T の数値を <see cref="CultureInfo.InvariantCulture"/> を使用してフォーマットし、単位サイズのサフィックスを付けた文字列に変換します。</summary>
      public static string UnitSizeToText<T>(T numberOfBytes)
      {
         // CultureInfo.CurrentCulture は地域アプレットで設定されたカルチャを使用する。

         return UnitSizeToText(numberOfBytes, CultureInfo.CurrentCulture);
      }


      /// <summary>型 T の数値を指定された <paramref name="cultureInfo"/> を使用してフォーマットし、単位サイズのサフィックスを付けた文字列に変換します。</summary>
      public static string UnitSizeToText<T>(T numberOfBytes, CultureInfo cultureInfo)
      {
         const int kb = 1024;
         var index = 0;

         var bytes = Convert.ToDouble(numberOfBytes, CultureInfo.InvariantCulture);

         if (bytes < 0)
         {
            bytes = 0;
         }

         else
         {
            while (bytes > kb && index < SizeFormats.Length - 1)
            {
               bytes /= kb;
               index++;
            }
         }


         // "512,00 B" の代わりに "512 B" を返す。

         return $"{bytes.ToString(index == 0 ? "0" : "0.##", cultureInfo)} {SizeFormats[index]}";
      }


      /// <summary>複数の値のハッシュコードを結合します。</summary>
      public static int CombineHashCodesOf<T1, T2>(T1 arg1, T2 arg2) => HashCode.Combine(arg1, arg2);

      /// <summary>複数の値のハッシュコードを結合します。</summary>
      public static int CombineHashCodesOf<T1, T2, T3>(T1 arg1, T2 arg2, T3 arg3) => HashCode.Combine(arg1, arg2, arg3);

      /// <summary>複数の値のハッシュコードを結合します。</summary>
      public static int CombineHashCodesOf<T1, T2, T3, T4>(T1 arg1, T2 arg2, T3 arg3, T4 arg4) => HashCode.Combine(arg1, arg2, arg3, arg4);
   }
}
