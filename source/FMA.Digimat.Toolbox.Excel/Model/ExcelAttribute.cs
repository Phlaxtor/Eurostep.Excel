namespace FMA.Digimat.Toolbox.Excel.Model
{
    public enum ExcelAttributeType
    {
        None = 0,
        CellStyle,
        ColumnStyle,
        HeaderStyle,
        MergeCellStyle,
    }

    public abstract class ExcelAttribute : Attribute
    {
        public abstract ExcelAttributeType AttributeType { get; }
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelCellStyleAttribute<T> : ExcelCellStyleBaseAttribute
        where T : ExcelCellStyle
    {
        public ExcelCellStyleAttribute() : base(typeof(T))
        {
        }
    }

    public abstract class ExcelCellStyleBaseAttribute : ExcelStyleAttribute
    {
        protected ExcelCellStyleBaseAttribute(Type styleType) : base(styleType)
        {
        }

        public override sealed ExcelAttributeType AttributeType => ExcelAttributeType.CellStyle;
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelColumnStyleAttribute<T> : ExcelColumnStyleBaseAttribute
        where T : ExcelCellStyle
    {
        public ExcelColumnStyleAttribute() : base(typeof(T))
        {
        }
    }

    public abstract class ExcelColumnStyleBaseAttribute : ExcelStyleAttribute
    {
        protected ExcelColumnStyleBaseAttribute(Type styleType) : base(styleType)
        {
        }

        public override sealed ExcelAttributeType AttributeType => ExcelAttributeType.ColumnStyle;
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelHeaderStyleAttribute<T> : ExcelHeaderStyleBaseAttribute
        where T : ExcelCellStyle
    {
        public ExcelHeaderStyleAttribute() : base(typeof(T))
        {
        }
    }

    public abstract class ExcelHeaderStyleBaseAttribute : ExcelStyleAttribute

    {
        protected ExcelHeaderStyleBaseAttribute(Type styleType) : base(styleType)
        {
        }

        public override sealed ExcelAttributeType AttributeType => ExcelAttributeType.HeaderStyle;
    }

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class ExcelMergeColumnsAttribute<T> : ExcelMergeColumnsBaseAttribute
        where T : ExcelCellStyle
    {
        public ExcelMergeColumnsAttribute(string title, uint count) : base(title, count, typeof(T))
        {
        }
    }

    public abstract class ExcelMergeColumnsBaseAttribute : ExcelStyleAttribute
    {
        protected ExcelMergeColumnsBaseAttribute(string title, uint count, Type styleType) : base(styleType)
        {
            Count = count;
            Title = title;
        }

        public override sealed ExcelAttributeType AttributeType => ExcelAttributeType.MergeCellStyle;

        public uint Count { get; }

        public string Title { get; }
    }

    public abstract class ExcelStyleAttribute : ExcelAttribute
    {
        protected ExcelStyleAttribute(Type styleType)
        {
            StyleType = styleType;
        }

        public Type StyleType { get; }

        public virtual ExcelCellStyle GetStyle()
        {
            return (ExcelCellStyle)Activator.CreateInstance(StyleType);
        }
    }
}