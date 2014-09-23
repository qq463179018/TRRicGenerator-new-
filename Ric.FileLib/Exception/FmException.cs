using System;

namespace Ric.FileLib.Exception
{
    [Serializable]
    public class FmException : System.Exception
    {
        public FmException()
        {
            
        }

        public FmException(string message) : base(message)
        {
            
        }

        public FmException(string message, System.Exception inner) : base(message, inner)
        {
            
        }

        protected FmException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
            
        }
    }
}
