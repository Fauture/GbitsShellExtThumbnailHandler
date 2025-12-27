#include "ByteArrayRead.h"
#pragma once
#include <iostream>
using namespace std;
// 带缓冲区初始化的构造函数
ByteArrayRead::ByteArrayRead(const void* buffer, uint32_t len)
	: _buffer(new char[len]), _pos(0), _length(len), _defaultEndian(ENDIAN_BIG) {
	memcpy(_buffer.get(), buffer, len);
}

ByteArrayRead::ByteArrayRead(ByteArrayRead&& other) noexcept
	: _buffer(std::move(other._buffer)),
	_pos(other._pos),
	_length(other._length),
	_defaultEndian(other._defaultEndian) {
	other._pos = 0;
	other._length = 0;
}

uint8_t ByteArrayRead::readByte() {
	return (_pos < _length) ? _buffer[_pos++] : 0;
}

uint16_t ByteArrayRead::readShort() {
	return readShort(_defaultEndian);
}

uint16_t ByteArrayRead::readShort(EndianType endian) {
	constexpr uint32_t SIZE = sizeof(uint16_t);
	if (_pos + SIZE > _length) return 0;

	uint16_t value;
	memcpy(&value, _buffer.get() + _pos, SIZE);
	_pos += SIZE;

	return (_systemEndian != _defaultEndian != endian) ? byteSwap(value) : value;
}

uint32_t ByteArrayRead::readInt() {
	return readInt(_defaultEndian);
}

uint32_t ByteArrayRead::readInt(EndianType endian) {
	constexpr uint32_t SIZE = sizeof(uint32_t);
	if (_pos + SIZE > _length) return 0;

	uint32_t value;
	memcpy(&value, _buffer.get() + _pos, SIZE);
	_pos += SIZE;

	return (_systemEndian != _defaultEndian != endian) ? byteSwap(value) : value;
}

// Varint32解码
uint32_t ByteArrayRead::readVarint32() {
	uint32_t result = 0;
	int shift = 0;
	uint8_t byte;

	do {
		if (_pos >= _length) return 0;
		byte = _buffer[_pos++];
		result |= (byte & 0x7F) << shift;
		shift += 7;
	} while (byte & 0x80);

	return result;
}

// Varint64解码
uint64_t ByteArrayRead::readVarint64() {
	uint64_t result = 0;
	int shift = 0;
	uint8_t byte;

	do {
		if (_pos >= _length) return 0;
		byte = _buffer[_pos++];
		result |= (static_cast<uint64_t>(byte & 0x7F)) << shift;
		shift += 7;
	} while (byte & 0x80);

	return result;
}

// ZigZag解码32位整数
int32_t ByteArrayRead::readZigZag32() {
	uint32_t value = readVarint32();
	// 使用条件表达式避免对无符号数直接取负
	return static_cast<int32_t>((value >> 1) ^ ((value & 1) ? ~0U : 0U));
}

// ZigZag解码64位整数
int64_t ByteArrayRead::readZigZag64() {
	uint64_t value = readVarint64();
	// 使用条件表达式避免对无符号数直接取负
	return static_cast<int64_t>((value >> 1) ^ ((value & 1) ? ~0ULL : 0ULL));
}

std::string ByteArrayRead::readStringByte() {
	const auto len = readByte();
	return readString(len);
}

std::string ByteArrayRead::readStringShort() {
	const auto len = readShort(_defaultEndian);
	return readString(len);
}
std::string ByteArrayRead::readStringShort(EndianType endian) {
	const auto len = readShort(endian);
	return readString(len);
}
std::string ByteArrayRead::readStringVarint32() {
	const auto len = readVarint32();
	return readString(len);
}

std::string ByteArrayRead::readString(uint32_t length) {
	if (length == 0 || _pos + length > _length)
		return "";

	std::string result(_buffer.get() + _pos, length);
	_pos += length;
	return result;
}

void ByteArrayRead::clear() {
	memset(_buffer.get(), 0, _length);
}

// 获取剩余数据
const char* ByteArrayRead::getRawBytes() {
	const int available = size();
	if (available <= 0) return nullptr;
	return _buffer.get() + _pos;
}

//获取指定长度数据
std::vector<char> ByteArrayRead::getRawBytes(uint32_t length) {
	length = min(length, size());
	std::vector<char> result(length, 0);
	if (length > 0) {
		memcpy(result.data(), _buffer.get() + _pos, length);
		_pos += length;
	}
	return result;
}

// 系统字节序检查（静态）
ByteArrayRead::EndianType ByteArrayRead::systemEndian() {
	union {
		uint32_t i;
		uint8_t c[4];
	} test = { 0x01020304 };
	return (test.c[0] == 0x01) ? ENDIAN_BIG : ENDIAN_LITTLE;
}

template <typename T>
T ByteArrayRead::byteSwap(T value) const {
	static_assert(std::is_integral<T>::value, "Only for integral types");
	return genericByteSwap(value);
}

template <typename T>
T ByteArrayRead::genericByteSwap(T value) const {
	union {
		T value;
		char bytes[sizeof(T)];
	} src, dst;

	src.value = value;
	for (uint32_t i = 0; i < sizeof(T); ++i) {
		dst.bytes[i] = src.bytes[sizeof(T) - 1 - i];
	}
	return dst.value;
}