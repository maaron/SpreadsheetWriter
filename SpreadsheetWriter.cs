using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;
using Functional;

namespace TMO_DPT
{
    public struct CellIndex
    {
        public int Row { get; private set; }
        public int Col { get; private set; }

        public CellIndex(int row, int col)
        {
            Row = row; Col = col;
        }
    }

    public struct CellSize
    {
        public int Height { get; private set; }
        public int Width { get; private set; }

        public CellSize(int height, int width)
        {
            Height = height; Width = width;
        }
    }

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

    public delegate CellSize Writer(SLDocument doc, CellIndex target);
    public delegate CellSize Writer<T>(SLDocument doc, CellIndex target, T value);

    public delegate void Formatter(SLDocument doc, CellRange target);
    public delegate void Formatter<T>(SLDocument doc, CellRange target, T value);

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

    public static class Writers
    {
        public static Writer<T> CellObject<T>()
        {
            return (doc, target, value) =>
            {
                doc.SetCellValue(
                    target.Row, target.Col, 
                    value.ToString());

                return new CellSize(1, 1);
            };
        }

        public static Writer<T> Cell<T>()
        {
            if (typeof(T) == typeof(string)) return CellString as Writer<T>;
            if (typeof(T) == typeof(int)) return CellInt as Writer<T>;
            if (typeof(T) == typeof(long)) return CellLong as Writer<T>;
            if (typeof(T) == typeof(double)) return CellDouble as Writer<T>;
            if (typeof(T) == typeof(DateTime)) return CellDateTime as Writer<T>;

            throw new NotSupportedException(String.Format(
                "Value of type {0} cannot be written to a cell", 
                typeof(T).Name));
        }

        public static Writer<Maybe<T>> MaybeCell<T>()
        {
            var cell = Cell<T>();
            return (doc, target, m) =>
            {
                if (m.HasValue) return cell(doc, target, m.Value);
                else return new CellSize(1, 1);
            };
        }

        public static Writer<string> CellString =
            (doc, target, value) =>
            {
                doc.SetCellValue(target.Row, target.Col, value);
                return new CellSize(1, 1);
            };

        public static Writer<int> CellInt =
            (doc, target, value) =>
            {
                doc.SetCellValue(target.Row, target.Col, value);
                return new CellSize(1, 1);
            };

        public static Writer<long> CellLong =
            (doc, target, value) =>
            {
                doc.SetCellValue(target.Row, target.Col, value);
                return new CellSize(1, 1);
            };

        public static Writer<double> CellDouble =
            (doc, target, value) =>
            {
                doc.SetCellValue(target.Row, target.Col, value);
                return new CellSize(1, 1);
            };

        public static Writer<DateTime> CellDateTime =
            (doc, target, value) =>
            {
                doc.SetCellValue(target.Row, target.Col, value);
                return new CellSize(1, 1);
            };

        public static Writer Const(string value)
        {
            return (doc, target) =>
            {
                doc.SetCellValue(
                    target.Row, target.Col, 
                    value);

                return new CellSize(1, 1);
            };
        }

        public static Writer<R> Select<T, R>(this Writer<T> writer, Func<R, T> selector)
        {
            return (doc, target, value) => writer(doc, target, selector(value));
        }

        public static Writer<R> Wrap<T, R>(this R t, Writer<T> writer, Func<R, T> selector)
        {
            return (doc, target, value) => writer(doc, target, selector(value));
        }

        public static Writer TopDown(this Writer a, Writer b)
        {
            return (doc, target) =>
            {
                var size = a(doc, target);
                return size.AddHeightMaxWidth(b(doc, target.Down(size)));
            };
        }

        public static Writer<T> TopDown<T>(this Writer a, Writer<T> b)
        {
            return (doc, target, value) =>
            {
                var size = a(doc, target);
                return size.AddHeightMaxWidth(b(doc, target.Down(size), value));
            };
        }

        public static Writer<IEnumerable<T>> TopDownMany<T>(this Writer<T> writer)
        {
            return (doc, target, source) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddHeightMaxWidth(writer(doc, target.Down(size), item)));
        }

        public static Writer TopDown<T>(this IEnumerable<T> source, Writer<T> writer)
        {
            return (doc, target) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddHeightMaxWidth(writer(doc, target.Down(size), item)));
        }

        public static Writer TopDown(this IEnumerable<Writer> writers)
        {
            return (doc, target) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddHeightMaxWidth(writer(doc, target.Down(size))));
        }

        public static Writer<T> TopDown<T>(this IEnumerable<Writer<T>> writers)
        {
            return (doc, target, value) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddHeightMaxWidth(writer(doc, target.Down(size), value)));
        }

        public static Writer<T> LeftRight<T>(this Writer a, Writer<T> b)
        {
            return (doc, target, value) =>
            {
                var size = a(doc, target);
                return size.AddWidthMaxHeight(b(doc, target.Right(size), value));
            };
        }

        public static Writer<IEnumerable<T>> LeftRightMany<T>(this Writer<T> writer)
        {
            return (doc, target, source) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddWidthMaxHeight(writer(doc, target.Right(size), item)));
        }

        public static Writer LeftRight<T>(this IEnumerable<T> source, Writer<T> writer)
        {
            return (doc, target) => source.Aggregate(
                new CellSize(0, 0),
                (size, item) => size.AddWidthMaxHeight(writer(doc, target.Right(size), item)));
        }

        public static Writer LeftRight(this IEnumerable<Writer> writers)
        {
            return (doc, target) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddWidthMaxHeight(writer(doc, target.Right(size))));
        }

        public static Writer<T> LeftRight<T>(this IEnumerable<Writer<T>> writers)
        {
            return (doc, target, value) => writers.Aggregate(
                new CellSize(0, 0),
                (size, writer) => size.AddWidthMaxHeight(writer(doc, target.Right(size), value)));
        }

        public static Writer Format(this Writer writer, Formatter formatter)
        {
            return (doc, target) =>
            {
                var size = writer(doc, target);
                formatter(doc, new CellRange(target, size));
                return size;
            };
        }

        public static Writer<T> Format<T>(this Writer<T> writer, Formatter formatter)
        {
            return (doc, target, value) =>
            {
                var size = writer(doc, target, value);
                formatter(doc, new CellRange(target, size));
                return size;
            };
        }
    }

    public static class Formatters
    {
        public static readonly Formatter Empty = delegate { };
    }

    public static class SLDocumentExtensions
    {
        public static void AddTable<T>(this SLDocument doc, int row, int col, Table<T> table)
        {
            table.Writer(doc, new CellIndex(row, col), table.Source);
        }
    }

    public static class Table
    {
        public static Table<T> Build<T>(this IEnumerable<T> source)
        {
            return new Table<T>(source);
        }
    }

    public class Table<T>
    {
        public IEnumerable<T> Source { get; private set; }

        public List<Tuple<Writer, Writer<T>>> Columns = new List<Tuple<Writer, Writer<T>>>();

        public Formatter RowFormatter { get; set; }
        public Formatter ColumnFormatter { get; set; }

        public Table(IEnumerable<T> source)
        {
            Source = source;
            RowFormatter = Formatters.Empty;
            ColumnFormatter = Formatters.Empty;
        }

        public Table<T> WithColumn<R>(string header, Func<T, R> selector)
        {
            return WithColumn(Writers.Const(header), selector, Writers.Cell<R>());
        }

        public Table<T> WithColumn<R>(string header, Func<T, Maybe<R>> selector)
        {
            return WithColumn(Writers.Const(header), selector, Writers.MaybeCell<R>());
        }

        public Table<T> WithColumn<R>(Writer header, Func<T, R> selector, Writer<R> writer)
        {
            Columns.Add(Tuple.Create(header, writer.Select(selector)));
            return this;
        }

        public Table<T> WithRowFormat(Formatter formatter)
        {
            RowFormatter = formatter;
            return this;
        }

        public Table<T> WithColumnFormat(Formatter formatter)
        {
            ColumnFormatter = formatter;
            return this;
        }

        public Writer<IEnumerable<T>> Writer
        {
            get
            {
                return HeaderWriter.TopDown(ContentWriter);
            }
        }

        public Writer HeaderWriter
        {
            get
            {
                return Columns.Select(c => c.Item1).LeftRight();
            }
        }

        public Writer<IEnumerable<T>> ContentWriter
        {
            get
            {
                return RowWriter.TopDownMany();
            }
        }

        public Writer<T> RowWriter
        {
            get
            {
                return Columns.Select(c => c.Item2).LeftRight();
            }
        }
    }

    public static class FooClass
    {
        public static void Foo(Writer<string> writer)
        {
            var source = new[] { 1, 2, 3 };

            var builder = Table.Build(source)
                .WithColumn("Value", i => i)
                .WithColumn(Writers.Const("SubColumn1"), i => i.ToString(), (d, r, s) => new CellSize(1, 1));
        }
    }
}
