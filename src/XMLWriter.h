#pragma once

#include "pugixml.hpp"

namespace dxfrg
{
	class xml_string_writer :
		public pugi::xml_writer
	{
	public:
		std::string Result;

		virtual void write( const void* data, size_t size )
		{
			Result.append( static_cast<const char*>(data), size );
		}
	};
}