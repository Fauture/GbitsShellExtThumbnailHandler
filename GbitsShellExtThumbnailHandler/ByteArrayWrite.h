#pragma once

#include <memory>
#include <string>
#include <stdexcept>
#include <type_traits>

class ByteArrayWrite {
public:
    enum EndianType {
        ENDIAN_BIG,
        ENDIAN_LITTLE
    };

    ByteArrayWrite(const ByteArrayWrite& other);

    explicit ByteArrayWrite(uint32_t initialCapacity = 256);
    ~ByteArrayWrite() = default;

    //写字节
    void writeByte(uint8_t value);
    //写短整数
    void writeShort( int16_t value);
    void writeShort(EndianType endian,int16_t value);
    //写整数
    void writeInt( int32_t value);
    void writeInt(EndianType endian,int32_t value);

    // Varint编码
    void writeVarint32(uint32_t value);
    void writeVarint64(uint64_t value);

    // ZigZag编码
    void writeZigZag32(int32_t value);
    void writeZigZag64(int64_t value);

    //写字节长度文本
    void writeStringByte(const std::string& str);
    //写短整数长度文本
    void writeStringShort(const std::string& str);
    void writeStringShort(EndianType endian, const std::string& str);

    // 字节集写入
    void writeBytes(const char* data, uint32_t len);

    // 内存访问接口
    const char* data() const noexcept;
    uint32_t size() const noexcept;
    void clear() noexcept;

    void copyFrom(const ByteArrayWrite& other);

private:

    std::unique_ptr<char[]> _buffer;
    uint32_t _writePos;
    uint32_t _capacity;
    EndianType _defaultEndian;
    EndianType _systemEndian = systemEndian();
    EndianType systemEndian();

    // 容量保障机制
    void ensureCapacity(uint32_t required);

    // 原始字节写入
    void writeRawBytes(const void* data, uint32_t len);

    // 基础类型序列化模板
    template <typename T>
    typename std::enable_if<std::is_integral<T>::value>::type
        writePrimitive(EndianType endian, T value);

    template <typename T>
    T byteSwap(T value) const;
    template <typename T>
    T genericByteSwap(T value) const;
};