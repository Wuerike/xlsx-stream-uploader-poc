-include .env
export

run: ## Run application
	dotnet watch run --project XlsxStreamUploader/XlsxStreamUploader.csproj
.PHONY: run

loadtest: ## Run load test benchmark (Release mode)
	dotnet run -c Release --project XlsxStreamUploaderLoadTest/XlsxStreamUploaderLoadTest.csproj
.PHONY: loadtest

loadtest-quick: ## Run quick validation test (Debug mode)
	dotnet run --project XlsxStreamUploaderLoadTest/XlsxStreamUploaderLoadTest.csproj
.PHONY: loadtest-quick