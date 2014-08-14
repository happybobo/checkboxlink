//
//  ExcelHandleViewController.m
//  CheckboxLink
//
//  Created by Yuichi Fujiki on 8/14/14.
//  Copyright (c) 2014 yfujiki. All rights reserved.
//

#import "ExcelHandleViewController.h"
#import <LibXL/LibXL.h>

@interface ExcelHandleViewController ()

- (IBAction)editXLS:(id)sender;
- (IBAction)editXLSX:(id)sender;
- (IBAction)editXLSM:(id)sender;

@property (weak, nonatomic) IBOutlet UILabel *outputLabel;

@property (nonatomic) BookHandle book;
@property (nonatomic) SheetHandle sheet;

@end

@implementation ExcelHandleViewController

- (id)initWithNibName:(NSString *)nibNameOrNil bundle:(NSBundle *)nibBundleOrNil
{
    self = [super initWithNibName:nibNameOrNil bundle:nibBundleOrNil];
    if (self) {
        // Custom initialization
    }
    return self;
}

- (void)viewDidLoad
{
    [super viewDidLoad];
    // Do any additional setup after loading the view.
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

/*
#pragma mark - Navigation

// In a storyboard-based application, you will often want to do a little preparation before navigation
- (void)prepareForSegue:(UIStoryboardSegue *)segue sender:(id)sender
{
    // Get the new view controller using [segue destinationViewController].
    // Pass the selected object to the new view controller.
}
*/

#pragma mark - XLS

- (NSString *)filePathForType:(NSString *)type {
    NSString *documentDir = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES)[0];
    NSString *outPath = [[documentDir stringByAppendingPathComponent:@"output"] stringByAppendingPathExtension:type];

    if ([[NSFileManager defaultManager] fileExistsAtPath:outPath]) {
        return outPath;
    }

    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"sample" ofType:type];
    return filePath;
}

- (void)readExcel:(NSString *)type {
    NSString *filePath = [self filePathForType:type];
    if ([[type lowercaseString] isEqualToString:@"xlsx"] || [[type lowercaseString] isEqualToString:@"xlsm"]) {
        self.book = xlCreateXMLBook();
    } else {
        self.book = xlCreateBook();
    }
    
    Boolean result = xlBookLoad(self.book, [filePath cStringUsingEncoding:NSUTF8StringEncoding]);

    if (!result) {
        NSLog(@"Failed to read sample.xlsx");
    }

    self.sheet = xlBookGetSheet(self.book, 0);
}

- (void)editExcel {
    [self toggleRow:32 column:1];
    [self toggleRow:33 column:2];
    [self toggleRow:34 column:1];
    [self toggleRow:34 column:2];
    [self toggleRow:34 column:3];
    [self toggleRow:35 column:2];
    [self toggleRow:35 column:3];
    [self toggleRow:35 column:6];

    xlSheetSetTopLeftView(self.sheet, 0, 0);

    //    for (int i = [self checkMarkRowMin]; i <= [self checkMarkRowMax]; i++) {
    //        xlSheetSetRowHidden(self.sheet, i, true);
    //    }
}

- (void)toggleRow:(NSInteger)row column:(NSInteger)column {
    Boolean val = xlSheetReadBool(self.sheet, row, column, 0);
    xlSheetWriteBool(self.sheet, row, column, !val, 0);
}

- (NSInteger)checkMarkRowMin {
    return 32;
}

- (NSInteger)checkMarkRowMax {
    return 35;
}

- (void)writeExcel:(NSString *)type {
    NSString *documentDir = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES)[0];
    NSString *outPath = [[documentDir stringByAppendingPathComponent:@"output"] stringByAppendingPathExtension:type];

    NSLog(@"Writing out excel file to %@", outPath);
    self.outputLabel.text = outPath;

    xlBookSave(self.book, [outPath cStringUsingEncoding:NSUTF8StringEncoding]);

    xlBookRelease(self.book);
}

- (char *)c_filePathFrom:(NSString *)filePath {

    char *c_filePath = malloc(256);
    [filePath getCString:c_filePath maxLength:256 encoding:NSUTF32BigEndianStringEncoding];
    return c_filePath;
}

- (IBAction)editXLS:(id)sender {
    [self readExcel:@"xls"];
    [self editExcel];
    [self writeExcel:@"xls"];
}

- (IBAction)editXLSX:(id)sender {
    [self readExcel:@"xlsx"];
    [self editExcel];
    [self writeExcel:@"xlsx"];
}

- (IBAction)editXLSM:(id)sender {
    [self readExcel:@"xlsm"];
    [self editExcel];
    [self writeExcel:@"xlsm"];
}
@end
