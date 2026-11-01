import Foundation
import CoreGraphics
import ImageIO

// Compiled with the unmodified flag-app TeamPlaybook.swift, never a second model.
@main struct AppSnapshotNative {
    static func require(_ condition: Bool, _ message: String) throws {
        if !condition { throw NSError(domain: "SnapshotContract", code: 1,
            userInfo: [NSLocalizedDescriptionKey: message]) }
    }
    static func main() throws {
        let args = CommandLine.arguments
        let root = URL(fileURLWithPath: args[2], isDirectory: true)
        if args[1] == "media" {
            let pixels = CGContext(data: nil, width: 320, height: 160, bitsPerComponent: 8,
                bytesPerRow: 1280, space: CGColorSpaceCreateDeviceRGB(),
                bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
            pixels.setFillColor(CGColor(red: 0.1, green: 0.5, blue: 0.2, alpha: 1))
            pixels.fill(CGRect(x: 0, y: 0, width: 320, height: 160))
            let png = NSMutableData()
            let destination = CGImageDestinationCreateWithData(png, "public.png" as CFString, 1, nil)!
            CGImageDestinationAddImage(destination, pixels.makeImage()!, nil)
            try require(CGImageDestinationFinalize(destination), "PNG generation failed")
            try (png as Data).write(to: root.appendingPathComponent("play.png"))
            let pdf = NSMutableData()
            var box = CGRect(x: 0, y: 0, width: 320, height: 160)
            let context = CGContext(consumer: CGDataConsumer(data: pdf)!, mediaBox: &box, nil)!
            context.beginPDFPage(nil); context.draw(pixels.makeImage()!, in: box)
            context.endPDFPage(); context.closePDF()
            try (pdf as Data).write(to: root.appendingPathComponent("coach.pdf"))
            return
        }
        let data = try Data(contentsOf: root.appendingPathComponent(args[3]))
        let team = args[4]
        if args[1] == "reject" {
            do {
                let library = try TeamPlaybookLibrary.decode(data, team: team)
                try verifyAssets(library, root)
            } catch { print("REJECTED"); return }
            throw NSError(domain: "SnapshotContract", code: 2,
                userInfo: [NSLocalizedDescriptionKey: "Unsafe fixture accepted"])
        }
        let library = try TeamPlaybookLibrary.decode(data, team: team)
        try verifyAssets(library, root)
        try require(library.books.count == 2, "Two books required")
        try require(library.defaultBook?.bookId == library.books[1].bookId, "Default star lost")
        let book = library.books[0]
        try require(book.plays.map(\.number) == [9, 4, 3], "Printed numbers/order changed")
        try require(book.plays.map(\.displayNumber) == ["9", "4", "C"], "Defense labels changed")
        try require(book.plays.map(\.name) == ["Cross · 左", "Out", "Zone"], "Names changed")
        try require(book.format == "mixed" && !book.coachesOnly && library.books[1].coachesOnly,
                    "Format or audience changed")
        try require(book.coachCards["offense"]?.playsPerPage == 9, "Coach layout changed")
        print("ACCEPTED")
    }
    static func verifyAssets(_ library: TeamPlaybookLibrary, _ root: URL) throws {
        for book in library.books {
            for asset in book.assets {
                let file = asset.mediaType == "image/png" ? "play.png" : "coach.pdf"
                try asset.verify(Data(contentsOf: root.appendingPathComponent(file)))
            }
        }
    }
}
