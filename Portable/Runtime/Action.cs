using System;
namespace OfficeExtension
{
    public class Action
    {
        private ActionInfo m_actionInfo;
		private bool m_isWriteOperation;

        public Action(ActionInfo actionInfo, bool isWriteOperation)
        {
            this.m_actionInfo = actionInfo;
            this.m_isWriteOperation = isWriteOperation;
        }

        internal ActionInfo ActionInfo
        {
            get
            {
                return this.m_actionInfo;
            }
		}

        internal bool IsWriteOperation
        { 
            get
            {
                return this.m_isWriteOperation;
            }
        }
	}
} 
