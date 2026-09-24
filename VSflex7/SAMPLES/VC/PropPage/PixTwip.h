// PixTwip.h: conversions between pixels and twips

#ifndef __PIXTWIP_H__
#define __PIXTWIP_H__

static _inline int GetDeviceConstants(BOOL x)
{
	static short lpx = 0;
	static short lpy = 0;
	if (!lpx)
	{
		HDC hdc = GetDC(NULL);
		lpx = GetDeviceCaps(hdc, LOGPIXELSX); 
		lpy = GetDeviceCaps(hdc, LOGPIXELSY);
		ReleaseDC(NULL, hdc);
	}
	return (x)? lpx: lpy;
}       

int _inline XTwipsToPixels(double twips)
{
	return MulDiv(GetDeviceConstants(true), (int)twips, 1440);
}

int _inline YTwipsToPixels(double twips)
{
	return MulDiv(GetDeviceConstants(false), (int)twips, 1440);
}  

long _inline XPixelsToTwips(int pixels)
{
	return MulDiv(pixels, 1440, GetDeviceConstants(true));
}

long _inline YPixelsToTwips(int pixels)
{
 	return MulDiv(pixels, 1440, GetDeviceConstants(false));
}  

#endif // __PIXTWIP_H__
