using System;
using System.Globalization;
using System.Reflection;
using WeihanLi.Extensions;

namespace WeihanLi.Npoi;

internal sealed class FakePropertyInfo : PropertyInfo
{
    private readonly Func<object?> _getValueFunc;
    private readonly object? _value;

    public FakePropertyInfo(Type entityType, Type propertyType, string propertyName)
    {
        DeclaringType = entityType;
        ReflectedType = entityType;
        PropertyType = propertyType;
        Name = propertyName;
        _value = propertyType.GetDefaultValue();
        _getValueFunc = () => _value;
        Attributes = PropertyAttributes.None;
    }

    public override Type DeclaringType { get; }
    public override string Name { get; }
    public override Type ReflectedType { get; }

    public override bool CanRead => false;
    public override bool CanWrite => false;
    public override Type PropertyType { get; }

    public override PropertyAttributes Attributes { get; }

    public override MethodInfo GetGetMethod(bool nonPublic) => _getValueFunc.Method;

    public override MethodInfo? GetSetMethod(bool nonPublic) => null;

    public override string ToString() => $"{PropertyType.Name}, {Name}";

    public override object[] GetCustomAttributes(bool inherit) => throw new NotSupportedException();

    public override object[] GetCustomAttributes(Type attributeType, bool inherit) =>
        throw new NotSupportedException();

    public override bool IsDefined(Type attributeType, bool inherit) => throw new NotSupportedException();

    public override MethodInfo[] GetAccessors(bool nonPublic) => throw new NotSupportedException();

    public override ParameterInfo[] GetIndexParameters() => throw new NotSupportedException();

    public override object? GetValue(object obj, BindingFlags invokeAttr, Binder binder, object[] index,
        CultureInfo culture) => _value;

    public override void SetValue(object obj, object? value, BindingFlags invokeAttr, Binder binder, object[] index,
        CultureInfo culture) => throw new NotSupportedException();
}
