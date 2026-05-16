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
 *
 * 
 * Copyright (c) Damien Guard.  All rights reserved.
 * AlphaFS has written permission from the author to include the CRC code.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;

namespace Alphaleonis.Win32.Security
{
   /// <summary>ISO-3309準拠の64ビットCRCハッシュアルゴリズムを実装します。</summary>
   [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "Crc")]
   internal class Crc64 : HashAlgorithm
   {
      private const ulong Iso3309Polynomial = 0xD800000000000000;
      private const ulong DefaultSeed = 0x0;


      private ulong m_hash;
      private readonly ulong m_seed;
      private readonly ulong[] m_table;
      private static ulong[] s_defaultTable;


      /// <summary><see cref="Crc64"/>の新しいインスタンスを初期化します。</summary>
      public Crc64() : this(Iso3309Polynomial, DefaultSeed)
      {
      }


      /// <summary><see cref="Crc64"/>の新しいインスタンスを初期化します。</summary>
      /// <param name="polynomial">多項式。</param>
      /// <param name="seed">シード値。</param>
      private Crc64(ulong polynomial, ulong seed)
      {
         m_table = InitializeTable(polynomial);
         m_seed = seed;
         m_hash = seed;
      }


      /// <summary><see cref="T:System.Security.Cryptography.HashAlgorithm"/>クラスの実装を初期化します。</summary>
      public override void Initialize()
      {
         m_hash = m_seed;
      }


      /// <summary>派生クラスでオーバーライドされた場合、オブジェクトに書き込まれたデータをハッシュ計算用のハッシュアルゴリズムにルーティングします。</summary>
      /// <param name="array">ハッシュコードを計算する入力データ。</param>
      /// <param name="ibStart">データの使用を開始するバイト配列内のオフセット。</param>
      /// <param name="cbSize">データとして使用するバイト配列内のバイト数。</param>
      protected override void HashCore(byte[] array, int ibStart, int cbSize)
      {
         m_hash = CalculateHash(m_hash, m_table, array, ibStart, cbSize);
      }


      /// <summary>暗号ストリームオブジェクトによって最後のデータが処理された後、ハッシュ計算を完了します。</summary>
      /// <returns>部分的な計算を完了し、データストリームの正しいハッシュ値を返します。</returns>
      protected override byte[] HashFinal()
      {
         var hashBuffer = UInt64ToBigEndianBytes(m_hash);
         HashValue = hashBuffer;
         return hashBuffer;
      }

      /// <summary>計算されたハッシュコードのサイズをビット単位で取得します。</summary>
      /// <value>計算されたハッシュコードのビット単位のサイズ。</value>
      public override int HashSize
      {
         get { return 64; }
      }


      /// <summary>テーブルを初期化します。</summary>
      /// <returns>初期化されたテーブル。</returns>
      /// <param name="polynomial">多項式。</param>
      private static ulong[] InitializeTable(ulong polynomial)
      {
         if (polynomial == Iso3309Polynomial && s_defaultTable != null)
         {
            return s_defaultTable;
         }

         var createTable = CreateTable(polynomial);

         if (polynomial == Iso3309Polynomial)
         {
            s_defaultTable = createTable;
         }

         return createTable;
      }


      /// <summary>テーブルを作成します。</summary>
      /// <returns>新しいulongの配列。</returns>
      /// <param name="polynomial">多項式。</param>
      private static ulong[] CreateTable(ulong polynomial)
      {
         var createTable = new ulong[256];

         for (var i = 0; i < 256; ++i)
         {
            var entry = (ulong)i;

            for (var j = 0; j < 8; ++j)
               entry = (entry & 1) == 1 ? (entry >> 1) ^ polynomial : entry >> 1;

            createTable[i] = entry;
         }

         return createTable;
      }


      /// <summary>ハッシュを計算します。</summary>
      /// <returns>計算されたハッシュ値。</returns>
      /// <param name="seed">シード値。</param>
      /// <param name="table">CRCテーブル。</param>
      /// <param name="buffer">入力バッファ。</param>
      /// <param name="start">開始位置。</param>
      /// <param name="size">サイズ。</param>
      private static ulong CalculateHash(ulong seed, ulong[] table, byte[] buffer, int start, int size)
      {
         // IList<byte> ではなく byte[] を直接受けることで JIT がインデクサ仮想ディスパッチを避け、
         // 境界チェック除去と直接アドレッシングが効くようになる。
         var hash = seed;

         for (var i = start; i < start + size; i++)
            unchecked
            {
               hash = (hash >> 8) ^ table[(buffer[i] ^ hash) & 0xff];
            }

         return hash;
      }


      /// <summary>UInt64値をビッグエンディアンのバイト配列に変換します。</summary>
      /// <returns>バイト配列。</returns>
      /// <param name="value">変換する値。</param>
      private static byte[] UInt64ToBigEndianBytes(ulong value)
      {
         return new byte[]
         {
            (byte)(value >> 56), (byte)(value >> 48), (byte)(value >> 40), (byte)(value >> 32),
            (byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value
         };
      }
   }
}
