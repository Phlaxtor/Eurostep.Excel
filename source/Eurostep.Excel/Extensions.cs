using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Net.Mail;

namespace Eurostep.Excel
{
    public static class Extensions
    {
        public static DateTime? FromExcel(this string? value)
        {
            if (string.IsNullOrEmpty(value)) return null;
            if (DateTime.TryParse(value, null, DateTimeStyles.AssumeUniversal, out DateTime v)) return v;
            throw new ApplicationException($"Can not parse date time value: '{value}'");
        }

        public static bool? GetBoolean(this string self)
        {
            bool boolValue;
            if (self.IsBoolean() && bool.TryParse(self, out boolValue))
            {
                return boolValue;
            }

            byte byteValue;
            if (self.IsByte() && byte.TryParse(self, out byteValue))
            {
                return Convert.ToBoolean(byteValue);
            }

            double doubleValue;
            if (self.IsDouble() && double.TryParse(self, out doubleValue))
            {
                return Convert.ToBoolean(doubleValue);
            }

            return null;
        }

        public static string? GetBooleanText(this string self)
        {
            var value = self.GetBoolean();
            if (value.HasValue) return value.Value.ToString();
            return null;
        }

        public static DateTime GetDate(this DateTime value)
        {
            if (value.TimeOfDay == TimeSpan.Zero) return value;
            DateTime result = value;
            if (value.Kind == DateTimeKind.Utc)
            {
                result = value.ToLocalTime();
            }
            if (value.Kind == DateTimeKind.Local)
            {
                result = value.ToUniversalTime();
            }
            return result.Date;
        }

        public static DateTime? GetDateTime(this string self)
        {
            DateTime dateValue;
            if (self.IsDateTime() && DateTime.TryParse(self, out dateValue))
            {
                return dateValue;
            }

            double doubleValue;
            if (self.IsDouble() && double.TryParse(self, out doubleValue))
            {
                return DateTime.FromOADate(doubleValue);
            }

            long longValue;
            if (self.IsLong() && long.TryParse(self, out longValue))
            {
                return DateTime.FromOADate(longValue);
            }

            return null;
        }

        public static string? GetDateTimeText(this string self)
        {
            var value = self.GetDateTime();
            if (value.HasValue) return value.Value.ToString();
            return null;
        }

        public static string? GetEmail(this string? s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            if (MailAddress.TryCreate(s, out MailAddress? _) == false) return null;
            return s;
        }

        public static string? GetId(this string? s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            if (int.TryParse(s, out int value)) return value.ToString("00000000");
            return s.Trim();
        }

        public static void GetId(this string? self, out string? value)
        {
            if (self.TryGetId(out value) == false) throw new ArgumentException("Id is not valid");
        }

        public static string? GetNumberText(this string self)
        {
            if (long.TryParse(self, out long l)) return l.ToString();
            if (double.TryParse(self, out double d)) return d.ToString();
            return null;
        }

        public static void GetValid(this string? self, out string? value)
        {
            if (self.TryGetValid(out value) == false) throw new ArgumentException("Value is not valid");
        }

        public static void GetVersion(this string? self, out string? value)
        {
            if (self.TryGetVersion(out value) == false) throw new ArgumentException("Version id is not valid");
        }

        public static string? GetVersionId(this string? s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            if (int.TryParse(s, out int value)) return value.ToString("000");
            string? fixedId = s.GetFixedVersionId();
            if (string.IsNullOrEmpty(fixedId)) return s;
            return fixedId;
        }

        public static bool IsBoolean(this string self)
        {
            if (!string.IsNullOrWhiteSpace(self))
            {
                bool x;
                if (bool.TryParse(self, out x))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsByte(this string self)
        {
            if (!string.IsNullOrWhiteSpace(self))
            {
                byte x;
                if (byte.TryParse(self, out x))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsDateTime(this string self)
        {
            if (!string.IsNullOrWhiteSpace(self))
            {
                DateTime x;
                if (DateTime.TryParse(self, out x))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsDouble(this string self)
        {
            if (!string.IsNullOrWhiteSpace(self))
            {
                double x;
                if (double.TryParse(self, out x))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsEqual(this string? self, string? other)
        {
            if (string.IsNullOrWhiteSpace(self)) return false;
            if (string.IsNullOrWhiteSpace(other)) return false;
            var s = self.Trim();
            var o = other.Trim();
            return string.CompareOrdinal(s, o) == 0;
        }

        public static bool IsLong(this string self)
        {
            if (!string.IsNullOrWhiteSpace(self))
            {
                long x;
                if (long.TryParse(self, out x))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool IsNumber(this string? value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (value.IsLong()) return true;
            if (value.IsDouble()) return true;
            return false;
        }

        public static bool IsValid(this string? self)
        {
            return self.TryGetValid(out _);
        }

        public static bool IsValidDate(this DateTime? self)
        {
            if (self.HasValue == false) return true;
            return self.Value.IsValidDate();
        }

        public static bool IsValidDate(this DateTime self)
        {
            var date = self.GetDate();
            if (date.TimeOfDay == TimeSpan.Zero) return true;
            return false;
        }

        public static bool IsValidEmail(this string? self)
        {
            return self.TryGetEmail(out _);
        }

        public static bool IsValidId(this string? self)
        {
            return self.TryGetId(out _);
        }

        public static bool IsValidVersion(this string? self)
        {
            return self.TryGetVersion(out _);
        }

        public static bool TryGetEmail(this string? self, [NotNullWhen(true)] out string? value)
        {
            string? s = self.GetEmail();
            return s.TryGetValid(out value);
        }

        public static bool TryGetId(this string? self, [NotNullWhen(true)] out string? value)
        {
            string? s = self.GetId();
            return s.TryGetValid(out value);
        }

        public static bool TryGetValid(this string? self, [NotNullWhen(true)] out string? value)
        {
            value = self;
            return string.IsNullOrEmpty(self) == false;
        }

        public static bool TryGetVersion(this string? self, [NotNullWhen(true)] out string? value)
        {
            string? s = self.GetVersionId();
            return s.TryGetValid(out value);
        }

        internal static string GetToUpperWithoutWhiteSpace(this string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            char[] result = new char[value.Length];
            int length = 0;
            for (int i = 0; i < value.Length; i++)
            {
                char s = value[i];
                switch (s)
                {
                    case ' ': continue;
                    case '\t': continue;
                    case '-': continue;
                    case '_': continue;
                    case '\u00A0': continue;
                    case '\0': continue;
                    default:
                        result[length++] = char.ToUpper(s);
                        break;
                }
            }
            return new string(result, 0, length);
        }

        private static string? GetFixedVersionId(this string s)
        {
            if (string.IsNullOrEmpty(s)) return null;
            string v = s.Substring(1);
            return v.GetVersionId();
        }
    }
}