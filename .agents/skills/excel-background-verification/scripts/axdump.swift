// Print the accessibility tree of an app's first window, one element per line,
// descending into WebKit web areas. System Events stops at the web area, so
// this is how to read rendered taskpane content when computer_use is unavailable.
//
//   swiftc -O -o /tmp/axdump axdump.swift && /tmp/axdump "Microsoft Excel" > ax.txt
//
// Needs Accessibility permission for the terminal running it. Read-only.
import Cocoa
import ApplicationServices

let args = CommandLine.arguments
let appName = args.count > 1 ? args[1] : "Microsoft Excel"
let maxDepth = args.count > 2 ? Int(args[2]) ?? 60 : 60
let maxElements = 20_000

guard let app = NSWorkspace.shared.runningApplications.first(where: { $0.localizedName == appName }) else {
    FileHandle.standardError.write("app not running: \(appName)\n".data(using: .utf8)!)
    exit(2)
}
let application = AXUIElementCreateApplication(app.processIdentifier)

func attribute(_ element: AXUIElement, _ name: String) -> AnyObject? {
    var value: AnyObject?
    return AXUIElementCopyAttributeValue(element, name as CFString, &value) == .success ? value : nil
}

func text(_ element: AXUIElement, _ name: String) -> String {
    guard let value = attribute(element, name) else { return "" }
    if let string = value as? String { return string }
    if let number = value as? NSNumber { return number.stringValue }
    if let url = value as? URL { return url.absoluteString }
    return ""
}

func children(_ element: AXUIElement) -> [AXUIElement] {
    attribute(element, kAXChildrenAttribute) as? [AXUIElement] ?? []
}

var printed = 0

func walk(_ element: AXUIElement, depth: Int) {
    if depth > maxDepth || printed >= maxElements { return }
    var parts = [text(element, kAXRoleAttribute)]
    let subrole = text(element, kAXSubroleAttribute)
    if !subrole.isEmpty { parts.append("(\(subrole))") }
    for (label, name) in [("title", kAXTitleAttribute), ("value", kAXValueAttribute), ("desc", kAXDescriptionAttribute), ("url", kAXURLAttribute)] {
        let value = text(element, name)
        if !value.isEmpty { parts.append("\(label)=\"\(value.prefix(200))\"") }
    }
    print(String(repeating: "  ", count: depth) + parts.joined(separator: " "))
    printed += 1
    for child in children(element) { walk(child, depth: depth + 1) }
}

guard let windows = attribute(application, kAXWindowsAttribute) as? [AXUIElement], let window = windows.first else {
    FileHandle.standardError.write("no windows for \(appName)\n".data(using: .utf8)!)
    exit(3)
}
walk(window, depth: 0)
