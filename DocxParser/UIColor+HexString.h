//
//  UIColor+HexString.h
//  DocxParser
//
//  Created by Bradley Slayter on 11/6/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <UIKit/UIKit.h>

@interface UIColor (HexString)

+(CGFloat) colorComponentFrom:(NSString *)string start:(NSUInteger)start length:(NSUInteger)length;
+(UIColor *) colorWithHexString:(NSString *)hexString;

@end
