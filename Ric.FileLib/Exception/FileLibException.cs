using System;

namespace Ric.FileLib.Exception
{
    [Serializable]
    public class FileLibException : System.Exception
    {
        public FileLibException()
        {
            
        }

        public FileLibException(string message) : base(message)
        {
            
        }

        public FileLibException(string message, System.Exception inner) : base(message, inner)
        {
            
        }

        protected FileLibException(
            System.Runtime.Serialization.SerializationInfo info,
            System.Runtime.Serialization.StreamingContext context)
            : base(info, context)
        {
            
        }
    }
}
