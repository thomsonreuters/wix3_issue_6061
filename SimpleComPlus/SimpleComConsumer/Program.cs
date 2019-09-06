using System;

namespace SimpleComConsumer
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("64-Bit Mode: {0}", Environment.Is64BitProcess);

			// Throwing nasty exceptions is good enough for this.
			int a = int.Parse(args[0]);
			int b = int.Parse(args[1]);

			scs.IDispCalculatorComponent calc = new scs.DispatchCalculatorComponent();

			int sum;
			calc.Add(a, b, out sum);
			Console.WriteLine("{0} + {1} = {2}", a, b, sum);
		}
	}
}
