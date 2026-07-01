#include "pch.h"
#include <memory>
#include "Log.h"

namespace MapiLibrary
{
	Log::Log(std::string loggerName, std::string logFilePath)
	{
		logger = Setup(loggerName, logFilePath);
	}

	std::shared_ptr<spdlog::logger> Log::Setup(
		std::string loggerName, std::string logFilePath)
	{
		std::shared_ptr<spdlog::logger> logger = spdlog::get(loggerName);

		if (not logger)
		{
			std::vector<spdlog::sink_ptr> sinks;

			std::shared_ptr<spdlog::sinks::stdout_sink_st> consoleLog =
				std::make_shared<spdlog::sinks::stdout_sink_st>();
			sinks.push_back(consoleLog);

			std::shared_ptr<spdlog::sinks::daily_file_sink_st> fileLog =
				std::make_shared<spdlog::sinks::daily_file_sink_st>(
					logFilePath, 23, 59);
			sinks.push_back(fileLog);

			logger = std::make_shared<spdlog::logger>(loggerName,
				std::begin(sinks),
				std::end(sinks));
			spdlog::register_logger(logger);
		}

		return logger;
	}
}
