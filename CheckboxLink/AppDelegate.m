//
//  AppDelegate.m
//  CheckboxLink
//
//  Created by Yuichi Fujiki on 2/7/14.
//  Copyright (c) 2014 yfujiki. All rights reserved.
//

#import "AppDelegate.h"
#import <LibXL/LibXL.h>

@interface AppDelegate()

@property (nonatomic) BookHandle book;
@property (nonatomic) SheetHandle sheet;

@end

@implementation AppDelegate

- (BOOL)application:(UIApplication *)application didFinishLaunchingWithOptions:(NSDictionary *)launchOptions
{
    self.window = [[UIWindow alloc] initWithFrame:[[UIScreen mainScreen] bounds]];

    self.window.backgroundColor = [UIColor whiteColor];
    [self.window makeKeyAndVisible];

    [self readXLS];
    [self editXLS];
    [self writeXLS];

    return YES;
}

- (void)applicationWillResignActive:(UIApplication *)application
{
    // Sent when the application is about to move from active to inactive state. This can occur for certain types of temporary interruptions (such as an incoming phone call or SMS message) or when the user quits the application and it begins the transition to the background state.
    // Use this method to pause ongoing tasks, disable timers, and throttle down OpenGL ES frame rates. Games should use this method to pause the game.
}

- (void)applicationDidEnterBackground:(UIApplication *)application
{
    // Use this method to release shared resources, save user data, invalidate timers, and store enough application state information to restore your application to its current state in case it is terminated later. 
    // If your application supports background execution, this method is called instead of applicationWillTerminate: when the user quits.
}

- (void)applicationWillEnterForeground:(UIApplication *)application
{
    // Called as part of the transition from the background to the inactive state; here you can undo many of the changes made on entering the background.
}

- (void)applicationDidBecomeActive:(UIApplication *)application
{
    // Restart any tasks that were paused (or not yet started) while the application was inactive. If the application was previously in the background, optionally refresh the user interface.
}

- (void)applicationWillTerminate:(UIApplication *)application
{
    // Called when the application is about to terminate. Save data if appropriate. See also applicationDidEnterBackground:.
}

#pragma mark - XLS

- (void)readXLS {
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"sample" ofType:@"xlsx"];
    self.book = xlCreateXMLBook();
    Boolean result = xlBookLoad(self.book, [filePath cStringUsingEncoding:NSUTF8StringEncoding]);

    if (!result) {
        NSLog(@"Failed to read sample.xlsx");
    }

    self.sheet = xlBookGetSheet(self.book, 0);
}

- (void)editXLS {
    xlSheetWriteBool(self.sheet, 32, 1, true, 0);
    xlSheetWriteBool(self.sheet, 33, 2, true, 0);
    xlSheetWriteBool(self.sheet, 34, 1, false, 0);
    xlSheetWriteBool(self.sheet, 34, 2, true, 0);
    xlSheetWriteBool(self.sheet, 34, 3, true, 0);
    xlSheetWriteBool(self.sheet, 35, 2, false, 0);
    xlSheetWriteBool(self.sheet, 35, 3, true, 0);
    xlSheetWriteBool(self.sheet, 35, 6, false, 0);

    xlSheetSetTopLeftView(self.sheet, 0, 0);

    for (int i = [self checkMarkRowMin]; i <= [self checkMarkRowMax]; i++) {
        xlSheetSetRowHidden(self.sheet, i, true);
    }
}

- (NSInteger)checkMarkRowMin {
    return 32;
}

- (NSInteger)checkMarkRowMax {
    return 35;
}

- (void)writeXLS {
    NSString *documentDir = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory, NSUserDomainMask, YES)[0];
    NSString *outPath = [documentDir stringByAppendingPathComponent:@"output.xlsx"];

    NSLog(@"Writing out excel file to %@", outPath);
    
    xlBookSave(self.book, [outPath cStringUsingEncoding:NSUTF8StringEncoding]);

    xlBookRelease(self.book);
}

- (char *)c_filePathFrom:(NSString *)filePath {

    char *c_filePath = malloc(256);
    [filePath getCString:c_filePath maxLength:256 encoding:NSUTF32BigEndianStringEncoding];
    return c_filePath;
}

@end
