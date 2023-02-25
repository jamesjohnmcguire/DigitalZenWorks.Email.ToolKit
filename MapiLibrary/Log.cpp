#include "pch.h"
#include <memory>
#include "Log.h"

namespace MapiLibrary
{
	Log::Log()
	{
		std::vector<spdlog::sink_ptr> sinks;

		std::shared_ptr<spdlog::sinks::stdout_sink_st> console_log =
			std::make_shared<spdlog::sinks::stdout_sink_st>();
		sinks.push_back(console_log);

		std::shared_ptr<spdlog::sinks::basic_file_sink_st> file_log =
			std::make_shared<spdlog::sinks::basic_file_sink_st>("logfile");
		sinks.push_back(file_log);

		logger = std::make_shared<spdlog::logger>(
			"log", begin(sinks), end(sinks));

		spdlog::set_pattern("%+");

		// spdlog::register_logger(logger);
	}

	std::shared_ptr<spdlog::logger> Log::Setup(
		std::string loggerName,
		std::vector<spdlog::sink_ptr> sinks)
	{
		auto logger = spdlog::get(loggerName);

		if (not logger)
		{
			if (sinks.size() > 0)
			{
				logger = std::make_shared<spdlog::logger>(loggerName,
					std::begin(sinks),
					std::end(sinks));
				spdlog::register_logger(logger);
			}
			else
			{
				logger = spdlog::stdout_color_mt(loggerName);
			}
		}

		return logger;
	}
}
