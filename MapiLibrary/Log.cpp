#include "pch.h"
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

		spdlog::register_logger(logger);
	}
}
