//
//  ViewController.m
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "ViewController.h"

@interface ViewController ()

@end

@implementation ViewController

- (void)viewDidLoad
{
    [super viewDidLoad];
	// Do any additional setup after loading the view, typically from a nib.
	
	NSURL *bundleURL = [[NSBundle mainBundle] URLForResource:@"Fable" withExtension:@"docx"];
	NSURL *localURL = [[[BSFileHelper sharedHelper] documentsDirectoryURL] URLByAppendingPathComponent:@"Fable.docx"];
	[[NSFileManager defaultManager] copyItemAtURL:bundleURL toURL:localURL error:nil];
	
	BSDocxRipperZipper *ripperZipper = [[BSDocxRipperZipper alloc] init];
	BSDocxParser *parser = [ripperZipper openDocxAtURL:localURL];
    [parser loadDocument];
    
    self.textView.attributedText = [parser getFinalString];
    self.textView.dataDetectorTypes = UIDataDetectorTypeLink;
	
	[ripperZipper writeStringToDocx:self.textView.attributedText];
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
