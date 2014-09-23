using System;

namespace Ric.FileLib.Exception
{
    [Serializable]
    public class IdnException : System.Exception
    {
        public IdnException()
        {
            
        }

        public IdnException(string message) : base(message)
        {
            
        }

        public IdnException(string message, System.Exception inner) : base(message, inner)
        {
            
        }

        protected IdnException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
            
        }
    }
}
