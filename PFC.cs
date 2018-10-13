/*
 * 由SharpDevelop创建。
 * 用户： ztf
 * 日期: 2016/6/9
 * 时间: 18:19
 * 
 * 要改变这种模板请点击 工具|选项|代码编写|编辑标准头文件
 */
using System;
using ExcelDna.Integration;
using System.Collections.Generic;

namespace PipeFrictionCoefficient
{
	/// <summary>
	/// Description of MyClass.
	/// </summary>
	public class PFC
	{
		[ExcelFunction(Description="Pipe Friction Coefficient")]
		/*	r is Reynolds number
		 * 	d2 is inner diameter of pipe,unit:mm
		 * 	e is Equivalent roughness
		 * 	q is Volume flow,unit:m^3/h
		 * 	v is specific volume,unit:m^3/kg
		*/
		public static string PFCcaculate(double r,double d2,double e,double q,double v)
		{
			double result = 0.0;
			string result1 = 0.ToString();
			double a1 = 26.98 * Math.Pow((d2/e),(double)(8/7));
			double a2 = 4160 * Math.Pow((d2/(2*e)),0.85);
			
			int caseSwitch = 1;
			if(r>0 && r<=2300)
			{
				caseSwitch = 1;
			}
			else if(r>2300 && r<=4000)
			{
				caseSwitch = 2;
			}
			else if(r>4000 && r<= a1)
			{
				caseSwitch = 3;
			}
			else if(r>a1 && r<=a2)
			{
				caseSwitch = 4;
			}
			else if(r>a2)
			{
				caseSwitch = 5;
			}
			
			switch (caseSwitch)
			{
				case 1:
					result = 64/r;
					break;
				case 2:
					result = -1;
					break;
				case 3:
					//result = Math.Pow((-3.924/(Math.Pow(r,2)) + 0.8/r),2);
					result = caculate1(r);
					break;
				case 4:
					result = 1.42*Math.Pow((Math.Log10(1.273*(q/(v*e)))) , -2);
					break;
				case 5:
					result = 1 / (Math.Pow((2*Math.Log10(d2/(2*e)))+1.74,2));
					break;
				default:
					result = 0;
					break;
			}
			
			if(result == -1.0)
			{
				result1 = "PICK DATA FROM PIC-7.3.3!";
			}
			else if(result == 0.0)
			{
				result1 = "DATA INVAILD!";
			}
			else
			{
				result1 = String.Format("{0:R}",result);
			}
			
			return result1;
		}
		
		private static double caculate1(double r)
		{
			int i = 1;
			double a = 0.0;
			double r1 = r;
			double m1 = 0.0;
			double m2 = 0.0;
			
			for(a = 0.001; a<=0.1; a+=0.001)
			{
				m1 = 1/Math.Pow(a,0.5);
				m2 = 2*Math.Log10(r1*Math.Pow(a,0.5)-0.8);
				if(Math.Abs((m1-m2)/m1)<0.01)
				{
					break;
				}
				i++;
			}
			
			if(i>=100)
			{
				return 0.0;
			}
			else
			{
				return a;
			}
		}
		
	}
}