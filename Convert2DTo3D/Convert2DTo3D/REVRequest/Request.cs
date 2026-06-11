namespace Convert2DTo3D
{
    public enum RequestId : int
    {
        None = -1,
        OK,
        Apply
    }

    public class Request
    {
        #region Variable

        private int m_request = (int)RequestId.None;

        #endregion Variable

        #region Method

        public RequestId Take()
        {
            return (RequestId)System.Threading.Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            System.Threading.Interlocked.Exchange(ref m_request, (int)request);
        }

        #endregion Method
    }
}