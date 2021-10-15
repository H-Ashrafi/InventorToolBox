using System;

namespace InventorToolBox.Tools
{
    public static class MathHelper
    {
        /// <summary>
        /// convert degrees to radian
        /// </summary>
        /// <param name="degree"></param>
        /// <returns></returns>
        public static double ToRadian(double degree)
        {
            return degree * Math.PI / 180;
        }
    }
}
