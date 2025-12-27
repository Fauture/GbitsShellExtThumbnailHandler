#include <string>
#include <cstring>
#include <vector>
#include <memory>
#include <Windows.h>
#pragma once

class ByteArrayRead
{
public:
	enum EndianType {
		ENDIAN_BIG,
		ENDIAN_LITTLE
	};

	//读字节
	uint8_t readByte();
	//读短整数
	uint16_t readShort();
	uint16_t readShort(EndianType endian);
	//读整数
	uint32_t readInt();
	uint32_t readInt(EndianType endian);

	// Varint解码
	uint32_t readVarint32();
	uint64_t readVarint64();

	// ZigZag解码
	int32_t readZigZag32();
	int64_t readZigZag64();

	//读字节长度文本
	std::string readStringByte();
	//读短整数长度文本
	std::string readStringShort();
	std::string readStringShort(EndianType endian);
	// Varint解码长度文本
	std::string readStringVarint32();
	//读指定长度文本//此文本默认无长度
	std::string readString(uint32_t length);

	// 获取原始数据
	const char* data() const noexcept { return _buffer.get(); }

	// 原始数据长度
	uint32_t length() const { return _length; }

	// 获取剩余数据长度
	uint32_t size() const { return _length - _pos; }

	// 获取/设置当前位置
	uint32_t position() const { return _pos; }
	void seek(uint32_t pos) {if (pos <= _length) _pos = pos;}

	//获取指定长度剩余数据
	std::vector<char> getRawBytes(uint32_t length);

	//获取剩余数据
	const char* getRawBytes();

	void clear();

	ByteArrayRead(ByteArrayRead&& other) noexcept;
	ByteArrayRead(const void* buffer, uint32_t len);

	ByteArrayRead(const ByteArrayRead&) = delete;
	ByteArrayRead& operator=(const ByteArrayRead&) = delete;

private:
	std::unique_ptr<char[]> _buffer;
	uint32_t _pos;
	uint32_t _length;
	EndianType _defaultEndian;
	EndianType _systemEndian = systemEndian();
	EndianType systemEndian();

	template <typename T>
	T byteSwap(T value) const;
	template <typename T>
	T genericByteSwap(T value) const;
};

