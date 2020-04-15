using System;
using System.Reflection;

namespace WeihanLi.Npoi
{
    internal class DelegateInvokeContext
    {
        public MethodInfo MethodInfo { get; }
        private object Target { get; }

        public DelegateInvokeContext(Delegate @delegate)
        {
            MethodInfo = @delegate.Method;
            Target = @delegate.Target;
        }

        public object Invoke(object[] parameters)
        {
            return MethodInfo?.Invoke(Target, parameters);
        }
    }
}
