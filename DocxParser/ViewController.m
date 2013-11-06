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
    
    BSDocxParser *parser = [[BSDocxParser alloc] initWithFileURL:[self getDocURL]];
    [parser loadDocument];
    
    self.textView.attributedText = [parser getFinalString];
}

-(NSURL *) getDocURL {
    return [[NSBundle mainBundle] URLForResource:@"document" withExtension:@"xml"];
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
