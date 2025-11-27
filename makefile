-include .env
export

run: ## Run application
	dotnet watch run --project XlsxStreamUploader/XlsxStreamUploader.csproj
.PHONY: run