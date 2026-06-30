#import <AppKit/AppKit.h>
#import <Foundation/Foundation.h>
#import <Vision/Vision.h>

static NSString *CleanField(NSString *value) {
    NSString *clean = [value stringByReplacingOccurrencesOfString:@"\t" withString:@" "];
    clean = [clean stringByReplacingOccurrencesOfString:@"\n" withString:@" "];
    clean = [clean stringByReplacingOccurrencesOfString:@"\r" withString:@" "];
    return clean;
}

int main(int argc, const char *argv[]) {
    @autoreleasepool {
        if (argc < 2) {
            fprintf(stderr, "usage: ocr_image <image-path>\n");
            return 2;
        }

        NSString *path = [NSString stringWithUTF8String:argv[1]];
        NSURL *imageURL = [NSURL fileURLWithPath:path];

        __block NSArray<VNRecognizedTextObservation *> *observations = @[];
        __block NSError *requestError = nil;
        VNRecognizeTextRequest *request =
            [[VNRecognizeTextRequest alloc] initWithCompletionHandler:^(VNRequest *finishedRequest, NSError *error) {
                if (error != nil) {
                    requestError = error;
                    return;
                }
                observations = (NSArray<VNRecognizedTextObservation *> *)finishedRequest.results;
            }];

        request.recognitionLevel = VNRequestTextRecognitionLevelAccurate;
        request.usesLanguageCorrection = YES;

        VNImageRequestHandler *handler = [[VNImageRequestHandler alloc] initWithURL:imageURL options:@{}];
        NSError *handlerError = nil;
        BOOL ok = [handler performRequests:@[ request ] error:&handlerError];
        if (!ok || handlerError != nil || requestError != nil) {
            NSError *error = handlerError ?: requestError;
            if (error != nil) {
                fprintf(stderr, "domain=%s code=%ld description=%s\n",
                        error.domain.UTF8String,
                        (long)error.code,
                        error.localizedDescription.UTF8String);
            } else {
                fprintf(stderr, "Vision request failed without NSError\n");
            }
            return 5;
        }

        for (VNRecognizedTextObservation *observation in observations) {
            VNRecognizedText *candidate = [[observation topCandidates:1] firstObject];
            if (candidate == nil || candidate.string.length == 0) {
                continue;
            }
            CGRect box = observation.boundingBox;
            NSString *text = CleanField(candidate.string);
            printf("%.6f\t%.6f\t%.6f\t%.6f\t%s\n",
                   box.origin.x,
                   box.origin.y,
                   box.size.width,
                   box.size.height,
                   text.UTF8String);
        }
    }
    return 0;
}
