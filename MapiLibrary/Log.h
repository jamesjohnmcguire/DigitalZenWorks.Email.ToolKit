#pragma once
#include <string>

#include "spdlog/spdlog.h"
#include "spdlog/sinks/basic_file_sink.h"
#include "spdlog/sinks/daily_file_sink.h"
#include "spdlog/sinks/stdout_sinks.h"
#include "spdlog/sinks/stdout_color_sinks.h"

#include "MapiLibrary.h"

namespace MapiLibrary
{
	class Log
	{
		public:
			Log();
			DllExport static std::shared_ptr<spdlog::logger> Setup(
				std::string loggerName,
				std::vector<spdlog::sink_ptr> sinks);

			template<typename T> void debug(const char*& message)
			{
				logger->debug(message);
			}

			template<typename T> void error(const char*& message)
			{
				logger->error(message);
			}

			template<typename T> void info(const T& message)
			{
				logger->info(message);
			}

			template<typename T> void warn(const char*& message)
			{
				logger->warn(message);
			}

	private:
			std::shared_ptr<spdlog::logger> logger;
	};
}
