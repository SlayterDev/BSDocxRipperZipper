//
//  ViewController.h
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <UIKit/UIKit.h>
#import "BSDocxParser.h"
#import "BSFileHelper.h"
#import "SSZipArchive.h"

@interface ViewController : UIViewController


@property (weak, nonatomic) IBOutlet UITextView *textView;

@end
