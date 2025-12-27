#pragma once
#include "ByteArrayWrite.h"


ByteArrayWrite::ByteArrayWrite(uint32_t initialCapacity)
	: _buffer(new char[initialCapacity]),
	_writePos(0),
	_capacity(initialCapacity),
	_defaultEndian(ENDIAN_BIG) {
	std::memset(_buffer.get(), 0, initialCapacity);
}

ByteArrayWrite::ByteArrayWrite(const ByteArrayWrite& other)
	: _buffer(new char[other._capacity]), _writePos(other._writePos), _capacity(other._capacity), _defaultEndian(other._defaultEndian)
{
	std::memcpy(_buffer.get(), other._buffer.get(), _capacity);
}

void ByteArrayWrite::writeByte(uint8_t value) {
	ensureCapacity(1);
	_buffer[_writePos++] = static_cast<char>(value);
}

void ByteArrayWrite::writeShort(int16_t value) {
	writePrimitive(_defaultEndian, value);
}

void ByteArrayWrite::writeShort(EndianType endian, int16_t value) {
	writePrimitive(endian, value);
}

void ByteArrayWrite::writeInt(int32_t value) {
	writePrimitive(_defaultEndian, value);
}

void ByteArrayWrite::writeInt(EndianType endian, int32_t value) {
	writePrimitive(endian, value);
}

// Varint32编码
void ByteArrayWrite::writeVarint32(uint32_t value) {
	while (value >= 0x80) {
		writeByte(static_cast<uint8_t>(value | 0x80));
		value >>= 7;
	}
	writeByte(static_cast<uint8_t>(value));
}

// Varint64编码
void ByteArrayWrite::writeVarint64(uint64_t value) {
	while (value >= 0x80) {
		writeByte(static_cast<uint8_t>(value | 0x80));
		value >>= 7;
	}
	writeByte(static_cast<uint8_t>(value));
}
// ZigZag32编码
void ByteArrayWrite::writeZigZag32(int32_t value) {
	// 使用无符号算术右移避免实现定义行为
	uint32_t uvalue = static_cast<uint32_t>((value << 1) ^ (value >> 31));
	writeVarint32(uvalue);
}

// ZigZag64编码
void ByteArrayWrite::writeZigZag64(int64_t value) {
	// 使用无符号算术右移避免实现定义行为
	uint64_t uvalue = static_cast<uint64_t>((value << 1) ^ (value >> 63));
	writeVarint64(uvalue);
}

void ByteArrayWrite::writeStringByte(const std::string& str) {
	if (str.size() > 0xFF)
		throw std::invalid_argument("String exceeds 1-byte length limit");
	writeByte(static_cast<uint8_t>(str.size()));
	writeRawBytes(str.data(), (uint32_t)str.size());
}

void ByteArrayWrite::writeStringShort(const std::string& str) {
	writeStringShort(_defaultEndian, str);
}

void ByteArrayWrite::writeStringShort(EndianType endian, const std::string& str) {
	if (str.size() > 0xFFFF)
		throw std::invalid_argument("String exceeds 2-byte length limit");
	writeShort(endian, static_cast<int16_t>(str.size()));
	writeRawBytes(str.data(), (uint32_t)str.size());
}

void ByteArrayWrite::writeBytes(const char* data, uint32_t len) {
	writeRawBytes(data, len);
}

const char* ByteArrayWrite::data() const noexcept {
	return _buffer.get();
}

uint32_t ByteArrayWrite::size() const noexcept {
	return _writePos;
}

void ByteArrayWrite::clear() noexcept {
	_writePos = 0;
	std::memset(_buffer.get(), 0, _capacity);
}

void ByteArrayWrite::ensureCapacity(uint32_t required) {
	if (_writePos + required <= _capacity) return;

	uint32_t newCapacity = (_capacity * 2 > _writePos + required)
		? _capacity * 2
		: _writePos + required + 16;

	auto newBuffer = std::unique_ptr<char[]>(new char[newCapacity]);
	std::memcpy(newBuffer.get(), _buffer.get(), _writePos);
	_buffer.swap(newBuffer);
	_capacity = newCapacity;
}

void ByteArrayWrite::writeRawBytes(const void* data, uint32_t len) {
	ensureCapacity(len);
	std::memcpy(_buffer.get() + _writePos, data, len);
	_writePos += len;
}

void ByteArrayWrite::copyFrom(const ByteArrayWrite& other) {
	// 确保当前缓冲区足够容纳other的数据
	ensureCapacity(other.size());
	// 复制数据
	std::memcpy(_buffer.get() + _writePos, other.data(), other.size());
	// 更新写入位置
	_writePos += other.size();
}

template <typename T>
typename std::enable_if<std::is_integral<T>::value>::type
ByteArrayWrite::writePrimitive(EndianType endian, T value) {
	if (endian != _defaultEndian != _systemEndian) {
		value = byteSwap(value);
	}
	writeRawBytes(&value, sizeof(value));
}

// 系统字节序检查（静态）
ByteArrayWrite::EndianType ByteArrayWrite::systemEndian() {
	union {
		uint32_t i;
		uint8_t c[4];
	} test = { 0x01020304 };
	return (test.c[0] == 0x01) ? ENDIAN_BIG : ENDIAN_LITTLE;
}

template <typename T>
T ByteArrayWrite::byteSwap(T value) const {
	static_assert(std::is_integral<T>::value, "Only for integral types");
	return genericByteSwap(value);
}

template <typename T>
T ByteArrayWrite::genericByteSwap(T value) const {
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