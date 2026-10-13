import Foundation
import Vision
import AppKit

guard CommandLine.arguments.count > 1 else {
    fputs("Usage: ocr <image-path>\n", stderr)
    exit(1)
}

let imagePath = CommandLine.arguments[1]
guard let image = NSImage(contentsOfFile: imagePath),
      let tiffData = image.tiffRepresentation,
      let bitmap = NSBitmapImageRep(data: tiffData),
      let cgImage = bitmap.cgImage else {
    fputs("Error: Could not load image \(imagePath)\n", stderr)
    exit(1)
}

var items: [[String: Any]] = []
let request = VNRecognizeTextRequest { request, error in
    if let error = error {
        fputs("OCR Error: \(error.localizedDescription)\n", stderr)
        return
    }
    guard let observations = request.results as? [VNRecognizedTextObservation] else { return }
    for observation in observations {
        if let top = observation.topCandidates(1).first {
            let b = observation.boundingBox
            items.append([
                "x": b.origin.x,
                "y": b.origin.y,
                "w": b.size.width,
                "h": b.size.height,
                "text": top.string
            ])
        }
    }
}
request.recognitionLevel = .accurate

do {
    try VNImageRequestHandler(cgImage: cgImage, options: [:]).perform([request])
    let rawText = items.map { $0["text"] as! String }.joined(separator: "\n")
    let output: [String: Any] = [
        "items": items,
        "raw_text": rawText
    ]
    let data = try JSONSerialization.data(withJSONObject: output)
    FileHandle.standardOutput.write(data)
} catch {
    fputs("Error: \(error.localizedDescription)\n", stderr)
    exit(1)
}
