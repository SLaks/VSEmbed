using System;
using System.Linq;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;

namespace VSThemeBrowser.VisualStudio {
	public class InterfaceShimProxy : RealProxy, IRemotingTypeInfo {
		private readonly object sourceObject;
		private readonly Type targetInterface;

		public InterfaceShimProxy(Type targetInterface, object sourceObject)
			: base(typeof(ContextBoundObject)) {
			this.sourceObject = sourceObject;
			this.targetInterface = targetInterface;
		}

		public override IMessage Invoke(IMessage msg) {
			var call = msg as IMethodCallMessage;
			if (msg == null)
				throw new NotSupportedException();

			var parameterTypes = call.MethodBase.GetParameters().Select(p => p.ParameterType).ToArray();
			var sourceMethod = sourceObject.GetType().GetMethod(call.MethodName, parameterTypes);

			var returnValue = sourceMethod.Invoke(sourceObject, call.Args);

			return new ReturnMessage(returnValue, call.Args, call.ArgCount, call.LogicalCallContext, call);
		}

		public bool CanCastTo(Type fromType, object o) {
			return fromType == targetInterface;
		}

		public string TypeName {
			get { return this.GetType().Name; }
			set { }
		}
	}
}
