using System;

namespace Ric.FileLib.Exception
{
    [Serializable]
    public class TcException : System.Exception
    {
        public TcException()
        {
            
        }

        public TcException(string message) : base(message)
        {
            
        }

        public TcException(string message, System.Exception inner) : base(message, inner)
        {
            
        }

        protected TcException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
            
        }
    }
}
