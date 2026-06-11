namespace Convert2DTo3D
{
    public enum ParameterWaringStatus
    {
        /// <summary>
        /// Không có lỗi nào xảy ra
        /// </summary>
        None,

        /// <summary>
        /// Paramter khác tên tên của group trong file share parameter.
        /// </summary>
        Different_GroupShareName,

        /// <summary>
        /// Parameter khác type
        /// </summary>
        Different_ParamType,

        /// <summary>
        /// Parameter khác tên group trong project paramter.
        /// </summary>
        Different_GroupProject,

        /// <summary>
        /// Parameter khác VaryByGroupInstance.
        /// </summary>
        Different_VaryByGroupInstance,

        /// <summary>
        /// Parameter khác categories.
        /// </summary>
        Different_Categories,

        /// <summary>
        /// Khác nhau về instance or type trong project paramter.
        /// </summary>
        Different_InstanceOrType,

        /// <summary>
        /// Các thông số đều thoả mãn.
        /// </summary>
        Equal
    }
}