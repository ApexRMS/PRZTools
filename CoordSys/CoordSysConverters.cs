﻿using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace NCC.PRZTools
{

    /// <summary>
    /// Returns the opposite or 'Not' of the input - True for False and False for True
    /// </summary>
    public class ReverseBoolConverter : IValueConverter
    {

        /// <summary>
        /// Convert True to False and False to True
        /// </summary>
        public Object Convert(Object value, Type targetType, Object parameter, CultureInfo culture)
        {
            return !(bool)value;
        }

        /// <summary>
        /// Not implemented
        /// </summary>
        public Object ConvertBack(Object value, Type targetType, Object parameter, CultureInfo culture)
        {
            throw new InvalidOperationException("Converter cannot convert back.");
        }

    }

    /// <summary>
    /// Converts the given bool to a Visibility
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {

        /// <summary>
        /// Convert True to Visible and False to Hidden
        /// </summary>
        public Object Convert(Object value, Type targetType, Object parameter, CultureInfo culture)
        {
            if (value == null)
                return Visibility.Hidden;
            bool val = (bool)value;
            return val ? Visibility.Visible : Visibility.Hidden;
        }

        /// <summary>
        /// Not implemented
        /// </summary>
        public Object ConvertBack(Object value, Type targetType, Object parameter, CultureInfo culture)
        {
            throw new InvalidOperationException("Converter cannot convert back.");
        }

    }
}