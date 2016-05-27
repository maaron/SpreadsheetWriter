using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Functional;

/*
 * This file contains a small "combinator" style library that can be used to
 * build "spreadsheet writers".  The basic idea is similar to building parser
 * combinators, in that we start with very simple, primitive "writers" and
 * combine them into larger and more sophisticated writers.
 *
 * All writers, whether they are simple or complex are functions that have one
 * of the following signatures:
 *
 *   CellSize (SLDocument, CellIndex)
 *   CellSize (SLDocument, CellIndex, T)
 *
 * The second signature has a generic type parameter 'T', which can be
 * anything.  The idea is that the first kind is a writer that always writes
 * the same thing into the document (unless dependent on some external state),
 * whereas the second type is one that writes a specified value.
 *
 * For handling formatting, or any 'post-processing' of a written range of
 * cells, there are also a couple of "formatter" types defined, although all
 * formatting could be done by a writer just the same.
 *
 *   void (SLDocument, CellRange)
 *   void (SLDocument, CellRange, T)
 *
 * The difference is much the same as before- either a formatter has everything
 * it needs to know, or it takes some input parameter that affects the
 * formatter's behavior.  Note that instead of a CellIndex, formatters take a
 * CellRange, indicating what range of cells the formatting should be applied
 * to.  The second type might be considered a "conditional formatter", although
 * this is not exactly the same as Microsoft Excel's conditional formatting
 * feature.
 *
 * The writers shown above can be assembled into more complicated writers using
 * combinators, such as the following:
 *
 *   Writer TopDown(Writer a, Writer b) - This creates a writer from two other
 *   writers that writes "a" then "b" just below it.
 *
 *   Writer LeftRight(Writer a, Writer b) - Same as above, but left-to-right.
 *
 *   Writer TopDownAll(IEnumerable<Writer> writers) - Creates a writer from a
 *   list of writers and writes all of them from top-to-bottom.
 *
 *   Writer<IEnumerable<T>> LeftRightMany(Writer<TDoc, T> writer) - Creates a writer
 *   that writes a list of values, given a writer that handles an single value.
 *
 * In addition to the combinators, there is a specialized class called
 * TableBuilder that allows the user to build a "table writer".  A table writer
 * is one that takes a list of items, and writes column headings and a row for
 * each item.  Generally, the column heading and subsequent items are written
 * on a single row, but the class is actually a little more abstract than than.
 */

namespace SpreadsheetWriter
{
    public interface ISheet
    {
        void WriteCell<T>(CellIndex index, T value);
    }

    /// <summary>
    /// This struct is used to represent the location of a single cell within a
    /// worksheet, identified by a (row, column) pair.
    /// </summary>
    public struct CellIndex
    {
        public int Row { get; private set; }
        public int Col { get; private set; }

        public CellIndex(int row, int col)
        {
            Row = row; Col = col;
        }
    }

    /// <summary>
    /// This struct is used to represent the size of a rectangular set of
    /// cells, in units of cells.
    /// </summary>
    public struct CellSize
    {
        public int Height { get; private set; }
        public int Width { get; private set; }

        public CellSize(int height, int width)
        {
            Height = height; Width = width;
        }
    }

    /// <summary>
    /// This struct represents a rectangular range of cells, given by a "top
    /// left" corner and a size.
    /// </summary>
    public struct CellRange
    {
        public CellIndex TopLeft { get; private set; }
        public CellSize Size { get; private set; }

        public CellRange(CellIndex topLeft, CellSize size)
        {
            TopLeft = topLeft;
            Size = size;
        }
    }

    /// <summary>
    /// A delegate type that represents any spreadsheet writer.  A writer takes a document and a 
    /// target cell, writes content, and returns the size of the content that was written.
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="target"></param>
    /// <returns></returns>
    public delegate CellSize Writer<TDoc>(TDoc doc, CellIndex target) where TDoc : ISheet;

    /// <summary>
    /// A delegate type that represents any spreadsheet writer that writes values of type T.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="doc"></param>
    /// <param name="target"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public delegate CellSize Writer<TDoc, T>(TDoc doc, CellIndex target, T value);

    public delegate void Formatter<TDoc>(TDoc doc, CellRange target);
    public delegate void Formatter<TDoc, T>(TDoc doc, CellRange target, T value);

    public static class CellSizeExtensions
    {
        public static CellSize Min(this CellSize a, CellSize b)
        {
            return new CellSize(
                Math.Min(a.Height, b.Height),
                Math.Min(a.Width, b.Width));
        }

        public static CellSize Add(this CellSize a, CellSize b)
        {
            return new CellSize(a.Height + b.Height, a.Width + b.Width);
        }

        public static CellSize AddHeight(this CellSize a, CellSize b)
        {
            return new CellSize(a.Height + b.Height, a.Width);
        }

        public static CellSize AddHeightMaxWidth(this CellSize a, CellSize b)
        {
            return new CellSize(
                a.Height + b.Height,
                Math.Max(a.Width, b.Width));
        }

        public static CellSize AddWidth(this CellSize a, CellSize b)
        {
            return new CellSize(a.Height, a.Width + b.Width);
        }

        public static CellSize AddWidthMaxHeight(this CellSize a, CellSize b)
        {
            return new CellSize(
                Math.Max(a.Height, b.Height),
                a.Width + b.Width);
        }
    }

    public static class CellIndexExtensions
    {
        public static CellIndex Down(this CellIndex index, CellSize size)
        {
            return new CellIndex(index.Row + size.Height, index.Col);
        }

        public static CellIndex Right(this CellIndex index, CellSize size)
        {
            return new CellIndex(index.Row, index.Col + size.Width);
        }
    }

    public static class Writers<TDoc>
       where TDoc : ISheet
    {
        public static Writer<TDoc, T> Cell<T>()
        {
            return (doc, target, value) =>
            {
                doc.WriteCell(target, value);
                return new CellSize(1, 1);
            };
        }

        public static Writer<TDoc> Const<T>(T value)
        {
            return Cell<T>().Bind(value);
        }

        /// <summary>
        /// Returns a <see cref="Writer{T}"/> that writes a <see cref="Maybe{T}"/> value into a single cell, 
        /// unless <see cref="Maybe{T}.HasValue"/> is false.  In that case, the cell is left blank.  See 
        /// <seealso cref="Cell{T}"/> for restrictions on the type parameter T.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static Writer<TDoc, Maybe<T>> MaybeCell<T>()
        {
            var cell = Cell<T>();
            return (doc, target, m) =>
            {
                if (m.HasValue) return cell(doc, target, m.Value);
                else return new CellSize(1, 1);
            };
        }
    }

    public static class Writers
    {
        // Careful, it may be tempting to replace MaybeCell with 
        // Cell<T>().Optional(), but the behavior isn't the same.  MaybeCell 
        // always returns a size of (1,1), even when no value is written, 
        // whereas Optional() returns a size of (0,0) if nothing is written.
        public static Writer<TDoc, Maybe<T>> Optional<TDoc, T>(this Writer<TDoc, T> writer)
        {
            return (doc, target, m) =>
            {
                if (m.HasValue) return writer(doc, target, m.Value);
                else return new CellSize(0, 0);
            };
        }

        /// <summary>
        /// Transforms a <see cref="Writer{T}"/> to a <see cref="Writer{R}"/> given a function from R to T.  
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <typeparam name="R"></typeparam>
        /// <param name="writer"></param>
        /// <param name="selector"></param>
        /// <remarks>
        /// Similar to other Select methods used as LINQ operators, but reversed in the sense that 
        /// T is an input type, rather than a return type.  As a result, this Select 
        /// implementation can't be used as a "select" clause in a LINQ expression, because the 
        /// type R cannot be inferred by the compiler.
        /// </remarks>
        /// <returns></returns>
        public static Writer<TDoc, R> Select<TDoc, T, R>(this Writer<TDoc, T> writer, Func<R, T> selector)
        {
            return (doc, target, value) => writer(doc, target, selector(value));
        }

        /// <summary>
        /// Turns a <see cref="Writer{T}"/> into a Writer, by "binding" the specified value to it.  In 
        /// this sense, the writer becomes one that writes a constant value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static Writer<TDoc> Bind<TDoc, T>(this Writer<TDoc, T> writer, T value)
            where TDoc : ISheet
        {
            return (doc, target) => writer(doc, target, value);
        }

        /// <summary>
        /// Creates a Writer that calls the supplied writers to write their content to the 
        /// worksheet in a top-down fashion.
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static Writer<TDoc> TopDown<TDoc>(this Writer<TDoc> a, Writer<TDoc> b)
            where TDoc : ISheet
        {
            return (doc, target) =>
            {
                var size = a(doc, target);
                return size.AddHeightMaxWidth(b(doc, target.Down(size)));
            };
        }

        /// <summary>
        /// Same as <see cref="TopDown(Writer, Writer)"/>, but allows the second writer to accept 
        /// a value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> TopDown<TDoc, T>(this Writer<TDoc> a, Writer<TDoc, T> b)
            where TDoc : ISheet
        {
            return (doc, target, value) =>
            {
                var size = a(doc, target);
                return size.AddHeightMaxWidth(b(doc, target.Down(size), value));
            };
        }

        /// <summary>
        /// Takes a <see cref="Writer{T}"/> and converts it to a new writer that writes an 
        /// IEnumerable&ltT&gt; by arranging them in a top-down fashion.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <returns></returns>
        public static Writer<TDoc, IEnumerable<T>> TopDownMany<TDoc, T>(this Writer<TDoc, T> writer)
            where TDoc : ISheet
        {
            return (doc, target, source) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddHeightMaxWidth(writer(doc, target.Down(size), item)));
        }

        /// <summary>
        /// Takes an sequence of writers and turns it into a new writer that calls each writer in 
        /// succession, such that the content is written ina top-down fashion.
        /// </summary>
        /// <param name="writers"></param>
        /// <returns></returns>
        public static Writer<TDoc> TopDownAll<TDoc>(this IEnumerable<Writer<TDoc>> writers)
            where TDoc : ISheet
        {
            return (doc, target) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddHeightMaxWidth(writer(doc, target.Down(size))));
        }

        /// <summary>
        /// Same as <seealso cref="TopDownAll(IEnumerable{Writer})"/>, but for values of type T.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writers"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> TopDownAll<TDoc, T>(this IEnumerable<Writer<TDoc, T>> writers)
        {
            return (doc, target, value) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddHeightMaxWidth(writer(doc, target.Down(size), value)));
        }

        /// <summary>
        /// Same as <seealso cref="TopDown{T}(Writer, Writer{T})"/>, but writes in a left-to-right 
        /// fashion.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> LeftRight<TDoc, T>(this Writer<TDoc> a, Writer<TDoc, T> b)
            where TDoc : ISheet
        {
            return (doc, target, value) =>
            {
                var size = a(doc, target);
                return size.AddWidthMaxHeight(b(doc, target.Right(size), value));
            };
        }

        /// <summary>
        /// Same as <seealso cref="TopDownMany{T}(Writer{T})"/>, but writes in a left-to-right 
        /// fashion.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <returns></returns>
        public static Writer<TDoc, IEnumerable<T>> LeftRightMany<TDoc, T>(this Writer<TDoc, T> writer)
        {
            return (doc, target, source) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddWidthMaxHeight(writer(doc, target.Right(size), item)));
        }

        /// <summary>
        /// Same as <seealso cref="TopDownAll(IEnumerable{Writer})"/>, but writes in a 
        /// left-to-right fashion.
        /// </summary>
        /// <param name="writers"></param>
        /// <returns></returns>
        public static Writer<TDoc> LeftRightAll<TDoc>(this IEnumerable<Writer<TDoc>> writers)
            where TDoc : ISheet
        {
            return (doc, target) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddWidthMaxHeight(writer(doc, target.Right(size))));
        }

        /// <summary>
        /// Same as <seealso cref="TopDownAll{T}(IEnumerable{Writer{T}})"/>, but writes in a 
        /// left-to-right fashion.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writers"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> LeftRightAll<TDoc, T>(this IEnumerable<Writer<TDoc, T>> writers)
        {
            return (doc, target, value) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddWidthMaxHeight(writer(doc, target.Right(size), value)));
        }

        /// <summary>
        /// Turns a Writer into a new writer that formats the content it writes using the supplied 
        /// Formatter.
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="formatter"></param>
        /// <returns></returns>
        public static Writer<TDoc> Format<TDoc>(this Writer<TDoc> writer, Formatter<TDoc> formatter)
            where TDoc : ISheet
        {
            return (doc, target) =>
            {
                var size = writer(doc, target);
                formatter(doc, new CellRange(target, size));
                return size;
            };
        }

        /// <summary>
        /// Turns a <see cref="Writer{T}"/> into a new writer that formats the content it writes using the 
        /// supplied Formatter.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <param name="formatter"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> Format<TDoc, T>(this Writer<TDoc, T> writer, Formatter<TDoc> formatter)
        {
            return (doc, target, value) =>
            {
                var size = writer(doc, target, value);
                formatter(doc, new CellRange(target, size));
                return size;
            };
        }

        /// <summary>
        /// Turns a <see cref="Writer{T}"/> into a new writer that formats the content it writes 
        /// using the supplied <see cref="Formatter{T}"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="writer"></param>
        /// <param name="formatter"></param>
        /// <returns></returns>
        public static Writer<TDoc, T> Format<TDoc, T>(this Writer<TDoc, T> writer, Formatter<TDoc, T> formatter)
        {
            return (doc, target, value) =>
            {
                var size = writer(doc, target, value);
                formatter(doc, new CellRange(target, size), value);
                return size;
            };
        }
    }

    public static class Formatters<TDoc>
    {
        public static readonly Formatter<TDoc> Empty = delegate { };
    }

    public static class Formatters<TDoc, T>
    {
        public static Formatter<TDoc, T> Empty = (doc, target, value) => { };
    }

    public static class Table
    {
        /// <summary>
        /// Creates a <see cref="TableBuilder{T}"/> from an <see cref="IEnumerable{T}"/>.  In 
        /// languages that support type-inference, this method is particularly useful when the 
        /// type <typeparamref name="T"/> is not easily expressed.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source"></param>
        /// <returns></returns>
        public static TableBuilder<TDoc, T> Build<TDoc, T>(TDoc doc, IEnumerable<T> source)
            where TDoc : ISheet
        {
            return new TableBuilder<TDoc, T>(source);
        }
    }

    /// <summary>
    /// This class can be used to build table writers, given a particular 
    /// <see cref="IEnumerable{T}"/> type.  The class is particularly useful when using a language 
    /// that supports type-inference.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <example>
    /// As an example, the following code creates a writer that writes a simple table to a 
    /// spreadsheet, given a list of string objects.
    /// <code>
    /// var source = new[] { "one", "two", "three" };
    /// 
    /// // Create the table writer
    /// var tableWriter = Table.Build(source)
    ///     .WithColumn("Length", s => s.Length)
    ///     .WithColumn("Value", s => s);
    ///     
    /// // Create a spreadsheet document
    /// var doc = new SLDocument();
    /// 
    /// // Write the table in the top left of the document
    /// tableWriter.Writer(doc, new CellIndex(1, 1), source);
    /// </code>
    /// </example>
    public class TableBuilder<TDoc, T>
        where TDoc : ISheet
    {
        public IEnumerable<T> Source { get; private set; }

        public List<Tuple<Writer<TDoc>, Writer<TDoc, T>>> Columns = new List<Tuple<Writer<TDoc>, Writer<TDoc, T>>>();

        public Formatter<TDoc, T> RowFormatter { get; set; }
        public Formatter<TDoc> HeaderFormatter { get; set; }
        public Formatter<TDoc> Formatter { get; set; }

        public TableBuilder(IEnumerable<T> source)
        {
            Source = source;
            RowFormatter = Formatters<TDoc, T>.Empty;
            HeaderFormatter = Formatters<TDoc>.Empty;
            Formatter = Formatters<TDoc>.Empty;
        }

        public TableBuilder<TDoc, T> WithColumn<R>(string header, Func<T, R> selector)
        {
            return WithColumn(Writers<TDoc>.Const(header), selector, Writers<TDoc>.Cell<R>());
        }

        public TableBuilder<TDoc, T> WithColumn<R>(string header, Func<T, Maybe<R>> selector)
        {
            return WithColumn(Writers<TDoc>.Const(header), selector, Writers<TDoc>.MaybeCell<R>());
        }

        public TableBuilder<TDoc, T> WithColumn<R>(string header, Func<T, R> selector, Writer<TDoc, R> writer)
        {
            return WithColumn(Writers<TDoc>.Const(header), selector, writer);
        }

        public TableBuilder<TDoc, T> WithColumn<R>(string header, Func<T, Maybe<R>> selector, Writer<TDoc, Maybe<R>> writer)
        {
            return WithColumn(Writers<TDoc>.Const(header), selector, writer);
        }

        public TableBuilder<TDoc, T> WithColumn<R>(Writer<TDoc> header, Func<T, R> selector, Writer<TDoc, R> writer)
        {
            Columns.Add(Tuple.Create(header, writer.Select(selector)));
            return this;
        }

        public TableBuilder<TDoc, T> WithFormat(Formatter<TDoc> formatter)
        {
            Formatter = formatter;
            return this;
        }

        public TableBuilder<TDoc, T> WithRowFormat(Formatter<TDoc, T> formatter)
        {
            RowFormatter = formatter;
            return this;
        }

        public Writer<TDoc, IEnumerable<T>> Writer
        {
            get
            {
                return HeaderWriter.TopDown(ContentWriter).Format(Formatter);
            }
        }

        public Writer<TDoc> HeaderWriter
        {
            get
            {
                return Columns.Select(c => c.Item1).LeftRightAll().Format(HeaderFormatter);
            }
        }

        public Writer<TDoc, IEnumerable<T>> ContentWriter
        {
            get
            {
                return RowWriter.TopDownMany();
            }
        }

        public Writer<TDoc, T> RowWriter
        {
            get
            {
                return Columns.Select(c => c.Item2).LeftRightAll().Format(RowFormatter);
            }
        }
    }
}
