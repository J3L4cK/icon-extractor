namespace IconExtractor {
	class Program {
		static int Main(string[] args) {
			var cli = new IconExtractorCLI();

			return cli.Execute(args);
		}
	}
}
