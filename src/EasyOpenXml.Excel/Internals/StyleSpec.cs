using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOpenXml.Excel.Internals
{

    #region _- FontSpec
    /// <summary>
    /// フォント仕様を表す
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct FontSpec : IEquatable<FontSpec>
    {
        /// <summary>
        /// フォント名を設定、取得します
        /// </summary>
        public readonly string Name;
        /// <summary>
        /// フォントサイズを設定、取得します
        /// </summary>
        public readonly double Size;
        /// <summary>
        /// フォント太字を設定、取得します
        /// </summary>
        public readonly bool Bold;
        /// <summary>
        /// フォントイタリックを設定、取得します
        /// </summary>
        public readonly bool Italic;
        /// <summary>
        /// フォント下線を設定、取得します
        /// </summary>
        public readonly bool Underline;
        /// <summary>
        /// フォント色を設定、取得します
        /// <para>フォント色を ARGB 16 進数文字列で指定します（例: "FFFF0000"）</para>
        /// <para>※指定されている場合、通常は Theme / Indexed より優先される</para>
        /// </summary>
        public readonly string Rgb; // "FFFF0000" など（ARGB）
        /// <summary>
        /// フォントテーマを設定、取得します
        /// </summary>
        public readonly int? Theme; // 使うなら
        /// <summary>
        /// フォント色を設定、取得します
        /// </summary>
        public readonly int? Indexed; // 使うなら

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="name">フォント名</param>
        /// <param name="size">フォントサイズ</param>
        /// <param name="bold">フォント太字</param>
        /// <param name="italic">イタリック設定（true：イタリック）</param>
        /// <param name="underline">下線設定（true：下線有り）</param>
        /// <param name="rgb">フォント色（ARGB 16 進数文字列で指定、例: "FFFF0000"）</param>
        /// <param name="theme">Excel のテーマカラーインデックス ※ rgb があればそれを優先</param>
        /// <param name="indexed">レガシーなインデックスカラーパレット番号 ※ rgb があればそれを優先</param>
        public FontSpec(string name, double size, bool bold, bool italic, bool underline, string rgb, int? theme, int? indexed)
        {
            Name = name ?? "";
            Size = size;
            Bold = bold;
            Italic = italic;
            Underline = underline;
            Rgb = rgb;       // null可にしてもOK、ここでは string のまま
            Theme = theme;
            Indexed = indexed;
        }

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public bool Equals(FontSpec other)
            => Name == other.Name && Size.Equals(other.Size) &&
               Bold == other.Bold && Italic == other.Italic && Underline == other.Underline &&
               Rgb == other.Rgb && Theme == other.Theme && Indexed == other.Indexed;

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="obj">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public override bool Equals(object obj)
            // 型パターンマッチング
            // A is Type B && Function(B)
            // A is Type であれば、(Type)A を B に代入し、Equals(other) を実施
            => obj is FontSpec other && Equals(other); 
        
        /// <summary>
        /// ハッシュ値を返します
        /// <para>Equals メソッドで用いる比較値は、ここで得られるハッシュ値となります</para>
        /// </summary>
        /// <returns>ハッシュ値</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + Name.GetHashCode();
                h = h * 31 + Size.GetHashCode();
                h = h * 31 + Bold.GetHashCode();
                h = h * 31 + Italic.GetHashCode();
                h = h * 31 + Underline.GetHashCode();
                h = h * 31 + (Rgb == null ? 0 : Rgb.GetHashCode());
                h = h * 31 + (Theme.HasValue ? Theme.Value.GetHashCode() : 0);
                h = h * 31 + (Indexed.HasValue ? Indexed.Value.GetHashCode() : 0);
                return h;
            }
        }
    }
    #endregion

    #region _- FillSpec
    /// <summary>
    /// 塗りつぶし仕様を表す
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct FillSpec : IEquatable<FillSpec>
    {
        /// <summary>
        /// 背景色を設定、取得します
        /// <para>フォント色を ARGB 16 進数文字列で指定します（例: "FFFF0000"）</para>
        /// </summary>
        public readonly string Rgb; // "FF00FF00" など。null/emptyなら「塗りなし」扱い可

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="rgb">背景色</param>
        public FillSpec(string rgb) { Rgb = rgb; }

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public bool Equals(FillSpec other) => Rgb == other.Rgb;

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public override bool Equals(object obj) => obj is FillSpec other && Equals(other);

        /// <summary>
        /// ハッシュ値を返します
        /// <para>Equals メソッドで用いる比較値は、ここで得られるハッシュ値となります</para>
        /// </summary>
        /// <returns>ハッシュ値</returns>
        public override int GetHashCode() => Rgb == null ? 0 : Rgb.GetHashCode();
    }
    #endregion

    #region _~ BorderLine
    /// <summary>
    /// 罫線列挙体
    /// </summary>
    internal enum BorderLine
    {
        None,
        Thin,
        Medium,
        Thick,
        Dashed,
        Dotted,
        Double
    }
    #endregion
    #region _~ BorderEdgeSpec
    /// <summary>
    /// 罫線仕様（1本単位）を表す
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct BorderEdgeSpec : IEquatable<BorderEdgeSpec>
    {
        /// <summary>
        /// 罫線種類
        /// </summary>
        public readonly BorderLine Line;
        /// <summary>
        /// 罫線の色
        /// </summary>
        public readonly string Rgb; // null可

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="line">罫線の種類</param>
        /// <param name="rgb">罫線の色</param>
        public BorderEdgeSpec(BorderLine line, string rgb) { Line = line; Rgb = rgb; }

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public bool Equals(BorderEdgeSpec other) => Line == other.Line && Rgb == other.Rgb;
        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="obj">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public override bool Equals(object obj) => obj is BorderEdgeSpec other && Equals(other);
        /// <summary>
        /// ハッシュ値を返します
        /// <para>Equals メソッドで用いる比較値は、ここで得られるハッシュ値となります</para>
        /// </summary>
        /// <returns>ハッシュ値</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                return ((int)Line * 397) ^ (Rgb == null ? 0 : Rgb.GetHashCode());
            }
        }
    }
    #endregion
    #region _- BorderSpec
    /// <summary>
    /// 罫線仕様を表す
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct BorderSpec : IEquatable<BorderSpec>
    {
        public readonly BorderEdgeSpec Left;
        public readonly BorderEdgeSpec Right;
        public readonly BorderEdgeSpec Top;
        public readonly BorderEdgeSpec Bottom;

        public BorderSpec(BorderEdgeSpec left, BorderEdgeSpec right, BorderEdgeSpec top, BorderEdgeSpec bottom)
        {
            Left = left; Right = right; Top = top; Bottom = bottom;
        }

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public bool Equals(BorderSpec other)
            => Left.Equals(other.Left) && Right.Equals(other.Right) && Top.Equals(other.Top) && Bottom.Equals(other.Bottom);

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="obj">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public override bool Equals(object obj) => obj is BorderSpec other && Equals(other);

        /// <summary>
        /// ハッシュ値を返します
        /// <para>Equals メソッドで用いる比較値は、ここで得られるハッシュ値となります</para>
        /// </summary>
        /// <returns>ハッシュ値</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + Left.GetHashCode();
                h = h * 31 + Right.GetHashCode();
                h = h * 31 + Top.GetHashCode();
                h = h * 31 + Bottom.GetHashCode();
                return h;
            }
        }
    }
    #endregion

    #region _- AllignmentSpec
    /// <summary>
    /// 配置仕様を表す
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct AlignmentSpec : IEquatable<AlignmentSpec>
    {

        /// <summary>
        /// 水平方向の配置情報
        /// </summary>
        public readonly DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues? H;
        /// <summary>
        /// 垂直方向の配置情報
        /// </summary>
        public readonly DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues? V;
        /// <summary>
        /// 折り返しフラグ（true: 折り返し）
        /// </summary>
        public readonly bool WrapText;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="h">水平方向の配置</param>
        /// <param name="v">垂直方向の配置</param>
        /// <param name="wrapText">折り返しフラグ</param>
        public AlignmentSpec(
            DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues? h,
            DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues? v,
            bool wrapText)
        {
            H = h;
            V = v;
            WrapText = wrapText;
        }

        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="other">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public bool Equals(AlignmentSpec other) => H == other.H && V == other.V && WrapText == other.WrapText;
        /// <summary>
        /// 指定されたオブジェクトと一致するかどうか比較します
        /// </summary>
        /// <param name="obj">比較対象</param>
        /// <returns>一致する場合は、true を返します。それ以外は、false</returns>
        public override bool Equals(object obj) => obj is AlignmentSpec other && Equals(other);
        /// <summary>
        /// ハッシュ値を返します
        /// <para>Equals メソッドで用いる比較値は、ここで得られるハッシュ値となります</para>
        /// </summary>
        /// <returns>ハッシュ値</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + (H.HasValue ? ((int)H.Value) : -1);
                h = h * 31 + (V.HasValue ? ((int)V.Value) : -1);
                h = h * 31 + WrapText.GetHashCode();
                return h;
            }
        }
    }

    #endregion

    #region _~ StyleKey
    /// <summary>
    /// スタイルキー（CellFormat の再利用キー）情報クラス
    /// <para>各スタイル仕様情報（FontSpec 等）を格納</para>
    /// <para>※1. 値オブジェクトのため、struct で実装</para>
    /// <para>※2. Dictionary の key とするため、IEquatable を実装</para>
    /// </summary>
    internal readonly struct StyleKey : IEquatable<StyleKey>
    {
        /// <summary>
        /// スタイルインデックス
        /// </summary>
        public readonly uint BaseStyleIndex;
        public readonly FontSpec? Font;
        public readonly FillSpec? Fill;
        public readonly BorderSpec? Border;
        public readonly AlignmentSpec? Alignment;

        public StyleKey(uint baseStyleIndex, FontSpec? font, FillSpec? fill, BorderSpec? border, AlignmentSpec? alignment)
        {
            BaseStyleIndex = baseStyleIndex;
            Font = font;
            Fill = fill;
            Border = border;
            Alignment = alignment;
        }

        public bool Equals(StyleKey other)
            => BaseStyleIndex == other.BaseStyleIndex &&
               NullableEquals(Font, other.Font) &&
               NullableEquals(Fill, other.Fill) &&
               NullableEquals(Border, other.Border) &&
               NullableEquals(Alignment, other.Alignment);

        private static bool NullableEquals<T>(T? a, T? b) where T : struct, IEquatable<T>
            => (!a.HasValue && !b.HasValue) || (a.HasValue && b.HasValue && a.Value.Equals(b.Value));

        public override bool Equals(object obj) => obj is StyleKey other && Equals(other);

        public override int GetHashCode()
        {
            unchecked
            {
                int h = 17;
                h = h * 31 + BaseStyleIndex.GetHashCode();
                h = h * 31 + (Font.HasValue ? Font.Value.GetHashCode() : 0);
                h = h * 31 + (Fill.HasValue ? Fill.Value.GetHashCode() : 0);
                h = h * 31 + (Border.HasValue ? Border.Value.GetHashCode() : 0);
                h = h * 31 + (Alignment.HasValue ? Alignment.Value.GetHashCode() : 0);
                return h;
            }
        }
    }

    #endregion
}
