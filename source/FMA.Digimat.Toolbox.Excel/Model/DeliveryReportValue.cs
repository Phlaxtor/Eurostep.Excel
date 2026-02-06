namespace FMA.Digimat.Toolbox.Excel.Model
{
    public sealed class DeliveryReportValue
    {
        public DeliveryReportValue()
        {
            IsUpdated = false;
            Value = string.Empty;
        }

        public DeliveryReportValue(string? value, bool isUpdated)
        {
            IsUpdated = isUpdated;
            Value = value ?? string.Empty;
        }

        public bool IsUpdated { get; }

        public string Value { get; }

        public static implicit operator bool(DeliveryReportValue value) => value.IsUpdated;

        public static implicit operator string(DeliveryReportValue value) => value.Value;

        public static bool operator false(DeliveryReportValue value) => value.IsUpdated == false;

        public static bool operator true(DeliveryReportValue value) => value.IsUpdated == true;

        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }

        public override string ToString()
        {
            return Value;
        }
    }
}