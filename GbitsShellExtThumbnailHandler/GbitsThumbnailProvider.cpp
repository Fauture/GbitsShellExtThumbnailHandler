#include "GbitsThumbnailProvider.h"
#include <Shlwapi.h>
#pragma comment(lib, "Shlwapi.lib")
#include<regex>
#include<string>
#include "webp/decode.h"
#pragma comment( lib, "lib/libwebp.lib" )
using namespace std;


//#define _DEBUG
#ifdef _DEBUG
#include<fstream>
#include <iostream>
#define DEBUGLOG(str) {std::wofstream file;file.open("C:\\log.txt", std::ios::app);file << str<< "\n";}
#else
#define DEBUGLOG(str) ;
#endif // DEBUG

extern HINSTANCE g_hInst;
extern long g_cDllRef;

bool RGBAtoBMP(byte* pImageData, int iWidth, int iHeight, UINT* pulSize, byte** ppOutData)
{
	if (pImageData == NULL ||
		iWidth == 0 ||
		iHeight == 0 ||
		pulSize == NULL ||
		ppOutData == NULL)
	{
		return false;
	}

	DWORD dwBytesPerPixel = 4;

	DWORD dwBMPWidth = (DWORD)iWidth;
	DWORD dwBMPHeight = (DWORD)iHeight;
	DWORD dwDataSize = dwBytesPerPixel * dwBMPWidth * dwBMPHeight;

	DWORD dwSize = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) + dwDataSize;

	char* pDataBuffer = (char*)malloc(dwSize);

	if (pDataBuffer == NULL)
	{
		return false;
	}

	// now setup the bitmap header

	BITMAPFILEHEADER* pBMFH = (BITMAPFILEHEADER*)pDataBuffer;
	pBMFH->bfType = ((USHORT)(BYTE)('B') | ((USHORT)(BYTE)('M') << 8));
	pBMFH->bfSize = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER) + dwDataSize;
	pBMFH->bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
	pBMFH->bfReserved1 = 0;
	pBMFH->bfReserved2 = 0;

	BITMAPINFOHEADER* pBMINFO = (BITMAPINFOHEADER*)(pDataBuffer + sizeof(BITMAPFILEHEADER));
	pBMINFO->biSize = sizeof(BITMAPINFOHEADER);
	pBMINFO->biWidth = dwBMPWidth;
	pBMINFO->biHeight = dwBMPHeight;
	pBMINFO->biCompression = 0;
	pBMINFO->biSizeImage = dwDataSize;
	pBMINFO->biXPelsPerMeter = 0;
	pBMINFO->biYPelsPerMeter = 0;
	pBMINFO->biClrUsed = 0;
	pBMINFO->biClrImportant = 0;

	pBMINFO->biPlanes = 1;
	pBMINFO->biBitCount = 32;

	// fill buffer data
	char* pDestData = pDataBuffer + sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);

	byte* pSrcData = pImageData;

	for (unsigned long ulRow = 0; ulRow < dwBMPHeight; ulRow++)
	{
		for (unsigned long ulCol = 0; ulCol < dwBMPWidth; ulCol++)
		{
			char cRed = *pSrcData++;
			char cGreen = *pSrcData++;
			char cBlue = *pSrcData++;
			char cAlpha = *pSrcData++;

			*pDestData++ = cBlue;
			*pDestData++ = cGreen;
			*pDestData++ = cRed;
			*pDestData++ = cAlpha;
		}
	} 

	*ppOutData = (byte*)pDataBuffer;
	*pulSize = dwSize;
	return true;
}

//解密微端gbits
bool Decrypt(byte* pImageData, int dwSize, UINT* pulSize, byte** ppOutData)
{
	if (pImageData == NULL ||
		dwSize == 0)
	{
		return false;
	}

	char* pDataBuffer = (char*)malloc(dwSize);

	if (pDataBuffer == NULL)
	{
		return false;
	}

	char* pDestData = pDataBuffer;
	byte* pSrcData = pImageData;


	int A = *pSrcData++;
	int B = *pSrcData++;
	int C = *pSrcData++;
	int D = *pSrcData++;
	int ret = ((D << 24) | (C << 16) | (B << 8) | (A << 0));

	//判断文件是否需要解密
	if (ret != 828732007) {
		*ppOutData = pImageData;
		*pulSize = dwSize;
		return true;
	}

	unsigned long dwBMPHeight = dwSize - 28;

	short cR = *pSrcData++;
	short cR1 = *pSrcData++;

	short cG = *pSrcData++;
	short cG1 = *pSrcData++;

	short cA = (cR << 8) | cR1;
	short cB = (cG << 8) | cG1;

	//校验文件 是否是加密文件
	if (cA != 22136 && cB != 39476) {
		return false;
	}

	for (unsigned long ulRow = 0; ulRow < 5; ulRow++)
	{
		short aR = *pSrcData++;
		short aR1 = *pSrcData++;

		short aG = *pSrcData++;
		short aG1 = *pSrcData++;

		short aA = (aR << 8) | aR1;
		short aB = (aG << 8) | aG1;

		short aC = aA ^ cA;
		short aD = aB ^ cB;

		char aC_1 = aC >> 8;;
		char aC_2 = (char)aC;

		char aD_1 = aD >> 8;;
		char aD_2 = (char)aD;

		*pDestData++ = aC_1;
		*pDestData++ = aC_2;
		*pDestData++ = aD_1;
		*pDestData++ = aD_2;
	}

	for (unsigned long ulRow = 0; ulRow < dwBMPHeight; ulRow++) {
		char R = *pSrcData++;
		*pDestData++ = R;
	}

	*ppOutData = (byte*)pDataBuffer;
	*pulSize = dwSize - 8;
	return true;
}

GbitsThumbnailProvider::GbitsThumbnailProvider() : m_cRef(1), m_pStream(NULL)
{
	InterlockedIncrement(&g_cDllRef);
}


GbitsThumbnailProvider::~GbitsThumbnailProvider()
{
	InterlockedDecrement(&g_cDllRef);
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP GbitsThumbnailProvider::QueryInterface(REFIID riid, void** ppv)
{
	static const QITAB qit[] =
	{
		QITABENT(GbitsThumbnailProvider, IThumbnailProvider),
		QITABENT(GbitsThumbnailProvider, IInitializeWithStream),
		{ 0 },
	};
	return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) GbitsThumbnailProvider::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) GbitsThumbnailProvider::Release()
{
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (0 == cRef)
	{
		delete this;
	}
	return cRef;
}

#pragma endregion


#pragma region IInitializeWithStream

// Initializes the thumbnail handler with a stream.
IFACEMETHODIMP GbitsThumbnailProvider::Initialize(IStream* pStream, DWORD grfMode)
{
	// A handler instance should be initialized only once in its lifetime. 
	HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);

	if (m_pStream == NULL)
	{
		// Take a reference to the stream if it has not been initialized yet.
		hr = pStream->QueryInterface(&m_pStream);
	}
	return hr;
}

#pragma endregion

//需要手动释放
byte* ReadUIntN(IStream* pStream, SIZE_T size)
{
	byte* imagex = (byte*)CoTaskMemAlloc(size);
	pStream->Read(imagex, size, NULL);
	return imagex;
}

//取出int
int ReadUInt32(IStream* pStream)
{
	unsigned char buf[4];
	pStream->Read(buf, 4, NULL);
	int ret = ((buf[3] << 24) | (buf[2] << 16) | (buf[1] << 8) | (buf[0] << 0));
	return ret;
}

//取出short
short ReadUInt16(IStream* pStream)
{
	unsigned char buf[2];
	pStream->Read(buf, 2, NULL);
	short ret = ((buf[1] << 8) | (buf[0] << 0));
	return ret;
}

//取出char
char ReadUInt8(IStream* pStream)
{
	unsigned char buf[1];
	pStream->Read(buf, 1, NULL);
	char ret = (buf[0] << 0);
	return ret;
}

//取出int
void ReadRun(IStream* pStream,int size)
{
	LARGE_INTEGER dlibMove;
	dlibMove.QuadPart = size;
	//pStream->Seek(dlibMove, STREAM_SEEK_SET, NULL);
	pStream->Seek(dlibMove, STREAM_SEEK_CUR, NULL);
}

//验证是否为图组
bool isZU(byte* bufferx)
{
	int A = *bufferx++;
	int B = *bufferx++;
	int C = *bufferx++;
	int D = *bufferx++;
	int ret = ((D << 24) | (C << 16) | (B << 8) | (A << 0));
	if (ret != 1234567890) {
		return false;
	}
	return true;
}

//查询图片位置并判断类型 (缩略图只需一张,所以直接搜索更直接)
int isType(byte* bufferx, int size, int* c)
{
	unsigned int dx = size / 2;

	int x = 0;
	for (unsigned int ulRow = 0; ulRow < dx; ulRow++)
	{
		int A = *bufferx++;
		int B = *bufferx++;
		int ret = ((B << 8) | (A << 0));
		if (ret == 40056) {//zlib压缩数据一般为Alpha,如果是特殊gbits图组则用于染色
			*c = x;
			return 0;
		}
		if (ret == 55551) {//jpg
			*c = x;
			return 1;
		}
		if (ret == 18770) {//webp
			*c = x;
			return 2;
		}
		x += 2;
	}
	return 0;
}


#pragma region IThumbnailProvider

// Gets a thumbnail image and alpha type. The GetThumbnail is called with the 
// largest desired size of the image, in pixels. Although the parameter is 
// called cx, this is used as the maximum size of both the x and y dimensions. 
// If the retrieved thumbnail is not square, then the longer axis is limited 
// by cx and the aspect ratio of the original image respected. On exit, 
// GetThumbnail provides a handle to the retrieved image. It also provides a 
// value that indicates the color format of the image and whether it has 
// valid alpha information.
IFACEMETHODIMP GbitsThumbnailProvider::GetThumbnail(UINT cx, HBITMAP* phbmp,
	WTS_ALPHATYPE* pdwAlpha)
{
	HRESULT hr;
	ULONG actRead;
	ULARGE_INTEGER fileSize;

	hr = IStream_Size(m_pStream, &fileSize); if (!SUCCEEDED(hr)) { DEBUGLOG("Error at IStream_Size"); return E_FAIL; }

	byte* buffer = (byte*)CoTaskMemAlloc(fileSize.QuadPart);
	if (!buffer) { DEBUGLOG("Error at malloc buffer");  return E_FAIL; }
	else
	{
		hr = m_pStream->Read(buffer, (LONG)fileSize.QuadPart + 1, &actRead);
		if (hr != S_FALSE)
		{
			CoTaskMemFree(buffer);
			DEBUGLOG("Error at Read Stream");
			return E_FAIL;
		}
	}

	byte* bufferx = (byte*)CoTaskMemAlloc(fileSize.QuadPart);
	UINT nFileLenx = 0;
	Decrypt((byte*)buffer, fileSize.QuadPart, &nFileLenx, &bufferx);

	if (isZU(bufferx))
	{
		DEBUGLOG("图组");
		int wz = 0;
		int type = isType(bufferx, nFileLenx, &wz);

		DEBUGLOG(type);
		DEBUGLOG(wz);
		if (type == 0) {
			CoTaskMemFree(buffer);
			return E_FAIL;
		}
		IStream* pImageStream = SHCreateMemStream(bufferx, nFileLenx);
		ReadRun(pImageStream, wz - 4);
		int datalen = ReadUInt32(pImageStream);
		byte* data = (byte*)CoTaskMemAlloc(datalen);
		data = ReadUIntN(pImageStream, datalen);
		DEBUGLOG("类型T");
		DEBUGLOG(datalen);
		if (type == 1) {//jpg
			IStream* pImageStreamx = SHCreateMemStream(data, datalen);
			hr = WICCreate32bppHBITMAP(pImageStreamx, phbmp, pdwAlpha);
			pImageStreamx->Release();
		}
		if (type == 2) {//webp
			int kx = 0;
			int gy = 0;
			WebPGetInfo(data, datalen, &kx, &gy);
			if (kx == 0 || gy == 0) {
				return E_FAIL;
			}
			byte* imagex = (byte*)CoTaskMemAlloc(static_cast<SIZE_T>(kx) * gy * 4);
			kx = 0;
			gy = 0;
			imagex = WebPDecodeRGBA(data, datalen, &kx, &gy);
			UINT size = 0;
			RGBAtoBMP((unsigned char*)imagex, kx, gy, &size, &imagex);

			IStream* pImageStreamx = SHCreateMemStream(imagex, size);
			hr = WICCreate32bppHBITMAP(pImageStreamx, phbmp, pdwAlpha);
			pImageStreamx->Release();
			CoTaskMemFree(imagex);
		}
		CoTaskMemFree(data);
		pImageStream->Release();
	}
	else
	{
		char A = *bufferx++;
		*bufferx--;
		IStream* pImageStream = SHCreateMemStream(bufferx, nFileLenx);
		ReadRun(pImageStream, 4);
		int datalen = ReadUInt32(pImageStream);
		byte* data = (byte*)CoTaskMemAlloc(datalen);
		data = ReadUIntN(pImageStream, datalen);
		int wz = 0;
		int type = isType(bufferx, nFileLenx, &wz);
		if (type == 0) {
			pImageStream->Release();
			CoTaskMemFree(buffer);
			return E_FAIL;
		}
		if (A == 8) {
			DEBUGLOG("webp单图");
			int kx = 0;
			int gy = 0;
			WebPGetInfo(data, datalen, &kx, &gy);
			if (kx==0 || gy==0) {
				return E_FAIL;
			}
			byte* imagex = (byte*)CoTaskMemAlloc(static_cast<SIZE_T>(kx) * gy * 4);
			kx = 0;
			gy = 0;
			imagex = WebPDecodeRGBA(data, datalen, &kx, &gy);
			UINT size = 0;
			RGBAtoBMP((unsigned char*)imagex, kx, gy, &size, &imagex);

			IStream* pImageStreamx = SHCreateMemStream(imagex, size);
			hr = WICCreate32bppHBITMAP(pImageStreamx, phbmp, pdwAlpha);
			pImageStreamx->Release();
			CoTaskMemFree(imagex);
		}
		else {
			DEBUGLOG("jpg单图");
			IStream* pImageStreamx = SHCreateMemStream(data, datalen);
			hr = WICCreate32bppHBITMAP(pImageStreamx, phbmp, pdwAlpha);
			pImageStreamx->Release();
		}
		CoTaskMemFree(data);
		pImageStream->Release();

	}
	CoTaskMemFree(buffer);
	return hr;
}

#pragma endregion


#pragma region Helper Functions

HRESULT GbitsThumbnailProvider::ConvertBitmapSourceTo32bppHBITMAP(
	IWICBitmapSource* pBitmapSource, IWICImagingFactory* pImagingFactory,
	HBITMAP* phbmp)
{
	*phbmp = NULL;

	IWICBitmapSource* pBitmapSourceConverted = NULL;
	WICPixelFormatGUID guidPixelFormatSource;
	HRESULT hr = pBitmapSource->GetPixelFormat(&guidPixelFormatSource);

	if (SUCCEEDED(hr) && (guidPixelFormatSource != GUID_WICPixelFormat32bppBGRA))
	{
		IWICFormatConverter* pFormatConverter;
		hr = pImagingFactory->CreateFormatConverter(&pFormatConverter);
		if (SUCCEEDED(hr))
		{
			// Create the appropriate pixel format converter.
			hr = pFormatConverter->Initialize(pBitmapSource,
				GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, NULL,
				0, WICBitmapPaletteTypeCustom);
			if (SUCCEEDED(hr))
			{
				hr = pFormatConverter->QueryInterface(&pBitmapSourceConverted);
			}
			pFormatConverter->Release();
		}
	}
	else
	{
		// No conversion is necessary.
		hr = pBitmapSource->QueryInterface(&pBitmapSourceConverted);
	}

	if (SUCCEEDED(hr))
	{
		UINT nWidth, nHeight;
		hr = pBitmapSourceConverted->GetSize(&nWidth, &nHeight);
		if (SUCCEEDED(hr))
		{
			BITMAPINFO bmi = { sizeof(bmi.bmiHeader) };
			bmi.bmiHeader.biWidth = nWidth;
			bmi.bmiHeader.biHeight = -static_cast<LONG>(nHeight);
			bmi.bmiHeader.biPlanes = 1;
			bmi.bmiHeader.biBitCount = 32;
			bmi.bmiHeader.biCompression = BI_RGB;

			BYTE* pBits;
			HBITMAP hbmp = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS,
				reinterpret_cast<void**>(&pBits), NULL, 0);
			hr = hbmp ? S_OK : E_OUTOFMEMORY;
			if (SUCCEEDED(hr))
			{
				WICRect rect = { 0, 0, (INT)nWidth, (INT)nHeight };

				// Convert the pixels and store them in the HBITMAP.  
				// Note: the name of the function is a little misleading - 
				// we're not doing any extraneous copying here.  CopyPixels 
				// is actually converting the image into the given buffer.
				hr = pBitmapSourceConverted->CopyPixels(&rect, nWidth * 4,
					nWidth * nHeight * 4, pBits);
				if (SUCCEEDED(hr))
				{
					*phbmp = hbmp;
				}
				else
				{
					DeleteObject(hbmp);
				}
			}
		}
		pBitmapSourceConverted->Release();
	}
	return hr;
}


HRESULT GbitsThumbnailProvider::WICCreate32bppHBITMAP(IStream* pstm,
	HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha)
{
	*phbmp = NULL;

	// Create the COM imaging factory.
	IWICImagingFactory* pImagingFactory;
	HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory, NULL,
		CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pImagingFactory));
	if (SUCCEEDED(hr))
	{
		// Create an appropriate decoder.
		IWICBitmapDecoder* pDecoder;
		hr = pImagingFactory->CreateDecoderFromStream(pstm,
			&GUID_VendorMicrosoft, WICDecodeMetadataCacheOnDemand, &pDecoder);
		if (SUCCEEDED(hr))
		{
			IWICBitmapFrameDecode* pBitmapFrameDecode;
			hr = pDecoder->GetFrame(0, &pBitmapFrameDecode);
			if (SUCCEEDED(hr))
			{
				hr = ConvertBitmapSourceTo32bppHBITMAP(pBitmapFrameDecode,
					pImagingFactory, phbmp);
				if (SUCCEEDED(hr))
				{
					*pdwAlpha = WTSAT_ARGB;
				}
				pBitmapFrameDecode->Release();
			}
			pDecoder->Release();
		}
		pImagingFactory->Release();
	}
	return hr;
}

#pragma endregion