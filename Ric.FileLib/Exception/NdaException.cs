using System;

namespace Ric.FileLib.Exception
{
    [Serializable]
    public class NdaException : System.Exception
    {
        public NdaException()
        {
            
        }

        public NdaException(string message) : base(message)
        {
            
        }

        public NdaException(string message, System.Exception inner) : base(message, inner)
        {
            
        }

        protected NdaException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
            
        }
    }
}
