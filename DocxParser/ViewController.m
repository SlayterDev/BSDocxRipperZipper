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
    self.textView.dataDetectorTypes = UIDataDetectorTypeLink;
	
	BSDocxWriter *writer = [[BSDocxWriter alloc] initWithAttributedString:self.textView.attributedText];
	NSData *data = [writer buildDocument];
	[self writeXML:data];
}

-(NSURL *) getDocURL {
    NSString *docxPath = [[NSBundle mainBundle] pathForResource:@"OP" ofType:@"docx"];
    NSString *zipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:@"OP.zip"];
    NSString *unZipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:@"OP"];
    NSError *error;
    [[NSFileManager defaultManager] moveItemAtPath:docxPath toPath:zipPath error:&error];
    if (error) {
        NSLog(@"Error: %@", error.localizedDescription);
    } else {
        NSLog(@"success");
    }
    
    [SSZipArchive unzipFileAtPath:zipPath toDestination:unZipPath];
    
    NSURL *docURL = [[BSFileHelper sharedHelper] documentsDirectoryURL];
    docURL = [docURL URLByAppendingPathComponent:@"OP/word/document.xml"];
    
    //return [[NSBundle mainBundle] URLForResource:@"document7" withExtension:@"xml"];
    return docURL;
}

-(void) fixOtherFile {
	NSURL *fileURL = [[BSFileHelper sharedHelper] documentsDirectoryURL];
	fileURL = [fileURL URLByAppendingPathComponent:@"OP/word/settings.xml"];
	
	NSData *fileData = [[NSData alloc] initWithContentsOfURL:fileURL];
	
	GDataXMLDocument *doc = [[GDataXMLDocument alloc] initWithData:fileData options:0 error:nil];
	GDataXMLElement *rsid = [[doc.rootElement elementsForName:@"w:rsids"] firstObject];
	
	[doc.rootElement removeChild:rsid];
	
	GDataXMLElement *newrsid = [GDataXMLElement elementWithName:@"w:rsids"];
	GDataXMLElement *rsidTag = [GDataXMLElement elementWithName:@"w:rsid"];
	GDataXMLNode *rsidNode = [GDataXMLNode elementWithName:@"w:val" stringValue:@"00000000"];
	[rsidTag addAttribute:rsidNode];
	GDataXMLElement *rsidRoot = [GDataXMLElement elementWithName:@"w:rsidRoot"];
	GDataXMLNode *rsidRootNode = [GDataXMLNode elementWithName:@"w:val" stringValue:@"00000000"];
	[rsidRoot addAttribute:rsidRootNode];
	[newrsid addChild:rsidRoot];
	[newrsid addChild:rsidTag];
	rsid = newrsid;
	
	[doc.rootElement addChild:rsid];
	
	GDataXMLDocument *newdoc = [[GDataXMLDocument alloc] initWithRootElement:doc.rootElement];
	[newdoc.XMLData writeToURL:fileURL atomically:YES];
}

-(void) writeXML:(NSData *)data {
	NSURL *filePath = [[BSFileHelper sharedHelper] documentsDirectoryURL];
	filePath = [filePath URLByAppendingPathComponent:@"here.xml"];
	[data writeToURL:filePath atomically:YES];
	
	[self fixOtherFile];
	
	NSLog(@"done %@", filePath);
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
