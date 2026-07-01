// pch.h: precompiled header file.

#ifndef PCH_H
#define PCH_H

#include <iostream>
#include <memory>
#include <string>
#include <vector>

#include "framework.h"

#include <MAPIDefS.h>
#include <EdkMdb.h>
#include <MAPITags.h>
#include <MAPIUtil.h>
#include <MAPIX.h>

#pragma warning( push )
#include "spdlog/spdlog.h"
#include "spdlog/sinks/basic_file_sink.h"
#include "spdlog/sinks/daily_file_sink.h"
#include "spdlog/sinks/stdout_sinks.h"
#include "spdlog/sinks/stdout_color_sinks.h"
#pragma warning(pop)

#endif //PCH_H
