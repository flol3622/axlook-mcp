// axlook-helper: a small, persistent Accessibility (AX) bridge for Microsoft Outlook (New Outlook for Mac).
//
// Protocol: one JSON object per line on stdin -> one JSON object per line on stdout.
//   request:  {"id": 1, "cmd": "find", ...args}
//   response: {"id": 1, "ok": true, "result": ...} | {"id": 1, "ok": false, "error": "..."}
//
// Elements are referred to by opaque handles ("e12"). Roots: "app", "main" (AXMainWindow), "focused"
// (focused element), "focusedWindow", "menubar". Handles stay valid until "release".
// All Outlook-specific logic lives in the TypeScript server; this helper only offers generic primitives.

import Cocoa
import ApplicationServices

let outlookBundleId = ProcessInfo.processInfo.environment["OUTLOOK_BUNDLE_ID"] ?? "com.microsoft.Outlook"
setvbuf(stdout, nil, _IOLBF, 0)

struct HelperError: Error { let message: String; init(_ m: String) { message = m } }

// MARK: - App & handles

var handles: [String: AXUIElement] = [:]
var handleCounter = 0
/// Extra attributes requested for every summary in the current request (e.g. AXDisclosureLevel).
var extraAttrs: [String] = []

func outlookApp() throws -> NSRunningApplication {
  // The running-application list can be transiently empty while it refreshes; retry briefly.
  for attempt in 0..<5 {
    if let app = NSRunningApplication.runningApplications(withBundleIdentifier: outlookBundleId).first(where: { !$0.isTerminated }) {
      return app
    }
    if attempt < 4 { usleep(100_000) }
  }
  throw HelperError("Microsoft Outlook is not running")
}

func appElement() throws -> AXUIElement {
  let el = AXUIElementCreateApplication(try outlookApp().processIdentifier)
  AXUIElementSetMessagingTimeout(el, 4)
  return el
}

func register(_ e: AXUIElement) -> String {
  handleCounter += 1
  let h = "e\(handleCounter)"
  handles[h] = e
  return h
}

func resolve(_ h: String) throws -> AXUIElement {
  let app = try appElement()
  switch h {
  case "app": return app
  case "main":
    // The Outlook main window (AXIdentifier "OutlookWindow"), even when a compose/event window is main.
    var candidates: [AXUIElement] = []
    if let w = attr(app, kAXMainWindowAttribute) { candidates.append(w as! AXUIElement) }
    if let w = attr(app, kAXFocusedWindowAttribute) { candidates.append(w as! AXUIElement) }
    candidates += (attr(app, kAXWindowsAttribute) as? [AXUIElement]) ?? []
    candidates += children(app).filter { str(attr($0, kAXRoleAttribute)) == "AXWindow" }
    if let w = candidates.first(where: { str(attr($0, "AXIdentifier")) == "OutlookWindow" }) { return w }
    guard let w = candidates.first, str(attr(w, kAXRoleAttribute)) == "AXWindow" else { throw HelperError("Outlook has no main window") }
    return w
  case "focusedWindow":
    guard let w = attr(app, kAXFocusedWindowAttribute) else { throw HelperError("Outlook has no focused window") }
    return w as! AXUIElement
  case "focused":
    guard let f = attr(app, kAXFocusedUIElementAttribute) else { throw HelperError("no focused element") }
    return f as! AXUIElement
  case "menubar":
    guard let m = attr(app, kAXMenuBarAttribute) else { throw HelperError("no menu bar") }
    return m as! AXUIElement
  default:
    guard let e = handles[h] else { throw HelperError("unknown handle \(h)") }
    return e
  }
}

// MARK: - Attribute helpers

func attr(_ e: AXUIElement, _ name: String) -> AnyObject? {
  var v: AnyObject?
  return AXUIElementCopyAttributeValue(e, name as CFString, &v) == .success ? v : nil
}

func str(_ o: AnyObject?) -> String? {
  guard let o = o else { return nil }
  if let s = o as? String { return s }
  if CFGetTypeID(o) == CFBooleanGetTypeID() { return (o as! Bool) ? "1" : "0" }
  if let n = o as? NSNumber { return n.stringValue }
  if CFGetTypeID(o) == CFDateGetTypeID() { return iso.string(from: o as! Date) }
  return nil
}

let iso: ISO8601DateFormatter = { let f = ISO8601DateFormatter(); f.formatOptions = [.withInternetDateTime]; return f }()

func children(_ e: AXUIElement) -> [AXUIElement] { (attr(e, kAXChildrenAttribute) as? [AXUIElement]) ?? [] }

func frame(_ e: AXUIElement) -> CGRect? {
  guard let v = attr(e, "AXFrame"), CFGetTypeID(v) == AXValueGetTypeID() else { return nil }
  var r = CGRect.zero
  return AXValueGetValue(v as! AXValue, .cgRect, &r) ? r : nil
}

/// JSON-friendly conversion of arbitrary AX attribute values.
func jsonValue(_ o: AnyObject?, depth: Int = 0) -> Any {
  guard let o = o else { return NSNull() }
  let t = CFGetTypeID(o)
  if t == AXUIElementGetTypeID() { return ["handle": register(o as! AXUIElement)] }
  if t == CFBooleanGetTypeID() { return o as! Bool }
  if let s = o as? String { return s }
  if let n = o as? NSNumber { return n }
  if t == CFDateGetTypeID() { return iso.string(from: o as! Date) }
  if t == AXValueGetTypeID() {
    let v = o as! AXValue
    switch AXValueGetType(v) {
    case .cgPoint: var p = CGPoint.zero; AXValueGetValue(v, .cgPoint, &p); return ["x": p.x, "y": p.y]
    case .cgSize: var s = CGSize.zero; AXValueGetValue(v, .cgSize, &s); return ["w": s.width, "h": s.height]
    case .cgRect: var r = CGRect.zero; AXValueGetValue(v, .cgRect, &r); return ["x": r.origin.x, "y": r.origin.y, "w": r.width, "h": r.height]
    case .cfRange: var r = CFRange(); AXValueGetValue(v, .cfRange, &r); return ["location": r.location, "length": r.length]
    default: return String(describing: v)
    }
  }
  if let a = o as? [AnyObject] { return depth > 1 ? a.count : a.prefix(200).map { jsonValue($0, depth: depth + 1) } }
  if let u = o as? URL { return u.absoluteString }
  return String(describing: o)
}

/// Compact summary of an element (only non-empty fields).
func summary(_ e: AXUIElement, withFrame: Bool = false, handle: String? = nil) -> [String: Any] {
  var d: [String: Any] = ["h": handle ?? register(e)]
  if let r = str(attr(e, kAXRoleAttribute)) { d["role"] = r }
  for (key, name) in [("subrole", kAXSubroleAttribute), ("id", "AXIdentifier"), ("title", kAXTitleAttribute),
                      ("desc", kAXDescriptionAttribute), ("help", kAXHelpAttribute)] {
    if let s = str(attr(e, name)), !s.isEmpty { d[key] = s }
  }
  if let v = attr(e, kAXValueAttribute) {
    if let s = str(v), !s.isEmpty { d["value"] = s }
  }
  if let sel = attr(e, kAXSelectedAttribute), CFGetTypeID(sel) == CFBooleanGetTypeID(), sel as! Bool { d["selected"] = true }
  if let en = attr(e, kAXEnabledAttribute), CFGetTypeID(en) == CFBooleanGetTypeID(), !(en as! Bool) { d["disabled"] = true }
  if withFrame, let f = frame(e) { d["frame"] = ["x": f.origin.x, "y": f.origin.y, "w": f.width, "h": f.height] }
  if !extraAttrs.isEmpty {
    var x: [String: Any] = [:]
    for n in extraAttrs { if let v = attr(e, n) { x[n] = jsonValue(v, depth: 2) } }
    if !x.isEmpty { d["x"] = x }
  }
  return d
}

func tree(_ e: AXUIElement, depth: Int, maxDepth: Int, maxChildren: Int, withFrame: Bool) -> [String: Any] {
  var d = summary(e, withFrame: withFrame)
  if depth < maxDepth {
    let kids = children(e)
    if !kids.isEmpty {
      d["children"] = kids.prefix(maxChildren).map { tree($0, depth: depth + 1, maxDepth: maxDepth, maxChildren: maxChildren, withFrame: withFrame) }
      if kids.count > maxChildren { d["more"] = kids.count - maxChildren }
    }
  } else if attr(e, kAXChildrenAttribute) != nil {
    let n = children(e).count
    if n > 0 { d["childCount"] = n }
  }
  return d
}

// MARK: - Matching

/// A matcher: every given key must match. String keys support exact match; *Prefix, *Contains and *Regex variants.
struct Matcher {
  let spec: [String: Any]
  func fieldValue(_ e: AXUIElement, _ field: String) -> String? {
    switch field {
    case "role": return str(attr(e, kAXRoleAttribute))
    case "subrole": return str(attr(e, kAXSubroleAttribute))
    case "id": return str(attr(e, "AXIdentifier"))
    case "title": return str(attr(e, kAXTitleAttribute))
    case "desc": return str(attr(e, kAXDescriptionAttribute))
    case "help": return str(attr(e, kAXHelpAttribute))
    case "value": return str(attr(e, kAXValueAttribute))
    case "any": // any of title/desc/value/help/id
      return [kAXTitleAttribute, kAXDescriptionAttribute, kAXValueAttribute, kAXHelpAttribute, "AXIdentifier"]
        .compactMap { str(attr(e, $0)) }.joined(separator: "\u{1F}")
    default: return nil
    }
  }
  func matches(_ e: AXUIElement) -> Bool {
    for (k, raw) in spec {
      guard let want = raw as? String else { continue }
      var field = k, mode = "eq"
      for suffix in ["Prefix", "Contains", "Regex", "Not"] where k.hasSuffix(suffix) && k.count > suffix.count {
        field = String(k.dropLast(suffix.count)); mode = suffix
      }
      let have = fieldValue(e, field) ?? ""
      switch mode {
      case "Prefix": if !have.hasPrefix(want) { return false }
      case "Contains": if !have.localizedCaseInsensitiveContains(want) { return false }
      case "Regex": if have.range(of: want, options: .regularExpression) == nil { return false }
      case "Not": if have == want { return false }
      default: if have != want { return false }
      }
    }
    return true
  }
}

func find(_ root: AXUIElement, _ m: Matcher, maxDepth: Int, limit: Int) -> [AXUIElement] {
  var out: [AXUIElement] = []
  var queue: [(AXUIElement, Int)] = [(root, 0)]
  var i = 0
  while i < queue.count && out.count < limit {
    let (e, d) = queue[i]; i += 1
    if d > 0 && m.matches(e) { out.append(e); if out.count >= limit { break } }
    if d < maxDepth { for k in children(e) { queue.append((k, d + 1)) } }
    if queue.count > 60000 { break } // safety valve for pathological trees
  }
  return out
}

// MARK: - Text extraction

let blockRoles: Set<String> = ["AXGroup", "AXParagraph", "AXList", "AXListItem", "AXRow", "AXHeading", "AXTable", "AXCell", "AXBlockquote", "AXWebArea", "AXSplitGroup", "AXScrollArea"]

/// Concatenate visible text below an element. Block-level containers are separated by newlines.
func collectText(_ e: AXUIElement, depth: Int = 0, links: inout [[String: String]]) -> String {
  if depth > 60 { return "" }
  let role = str(attr(e, kAXRoleAttribute)) ?? ""
  if role == "AXStaticText" { return str(attr(e, kAXValueAttribute)) ?? "" }
  if role == "AXTextArea" || role == "AXTextField" { return str(attr(e, kAXValueAttribute)) ?? "" }
  if role == "AXImage" {
    let d = str(attr(e, kAXDescriptionAttribute)) ?? ""
    return d.isEmpty ? "" : "[image: \(d)]"
  }
  if role == "AXLink", let url = attr(e, "AXURL") as? URL {
    links.append(["text": str(attr(e, kAXTitleAttribute)) ?? "", "url": url.absoluteString])
  }
  var parts: [String] = []
  for k in children(e) {
    let t = collectText(k, depth: depth + 1, links: &links)
    if !t.isEmpty { parts.append(t) }
  }
  let sep = blockRoles.contains(role) ? "\n" : ""
  return parts.joined(separator: sep)
}

/// WebKit exposes the rendered text of a web area through text markers; try that first.
func webAreaText(_ e: AXUIElement) -> String? {
  guard let range = attr(e, "AXTextMarkerRangeForUIElement") ?? paramAttr(e, "AXTextMarkerRangeForUIElement", e) else { return nil }
  return paramAttr(e, "AXStringForTextMarkerRange", range) as? String
}

func paramAttr(_ e: AXUIElement, _ name: String, _ param: AnyObject) -> AnyObject? {
  var v: AnyObject?
  return AXUIElementCopyParameterizedAttributeValue(e, name as CFString, param, &v) == .success ? v : nil
}

// MARK: - Input

func modifierFlags(_ mods: [String]) -> CGEventFlags {
  var f: CGEventFlags = []
  for m in mods {
    switch m.lowercased() {
    case "cmd", "command": f.insert(.maskCommand)
    case "shift": f.insert(.maskShift)
    case "opt", "option", "alt": f.insert(.maskAlternate)
    case "ctrl", "control": f.insert(.maskControl)
    default: break
    }
  }
  return f
}

/// Post a key press. Events go to the HID stream when Outlook is frontmost (reaches out-of-process panels),
/// otherwise directly to Outlook's pid.
func postKey(_ code: CGKeyCode, _ flags: CGEventFlags, pid: pid_t, global: Bool) {
  let src = CGEventSource(stateID: .hidSystemState)
  for down in [true, false] {
    let ev = CGEvent(keyboardEventSource: src, virtualKey: code, keyDown: down)!
    ev.flags = flags
    if global { ev.post(tap: .cghidEventTap) } else { ev.postToPid(pid) }
    usleep(8000)
  }
}

func typeText(_ text: String, pid: pid_t) throws {
  let src = CGEventSource(stateID: .hidSystemState)
  let units = Array(text.utf16)
  var i = 0
  while i < units.count {
    try requireFrontmost("keystrokes")
    var chunk = Array(units[i..<min(i + 16, units.count)])
    // never split a surrogate pair
    if let last = chunk.last, UTF16.isLeadSurrogate(last), i + chunk.count < units.count { chunk.append(units[i + chunk.count]) }
    for down in [true, false] {
      let ev = CGEvent(keyboardEventSource: src, virtualKey: 0, keyDown: down)!
      ev.keyboardSetUnicodeString(stringLength: chunk.count, unicodeString: &chunk)
      // Delivered to Outlook's process only: keystrokes can never reach another app.
      ev.postToPid(pid)
    }
    usleep(12000)
    i += chunk.count
  }
}

func click(at p: CGPoint, double: Bool = false) {
  let src = CGEventSource(stateID: .hidSystemState)
  let move = CGEvent(mouseEventSource: src, mouseType: .mouseMoved, mouseCursorPosition: p, mouseButton: .left)!
  move.post(tap: .cghidEventTap)
  usleep(30000)
  for n in 1...(double ? 2 : 1) {
    let down = CGEvent(mouseEventSource: src, mouseType: .leftMouseDown, mouseCursorPosition: p, mouseButton: .left)!
    down.setIntegerValueField(.mouseEventClickState, value: Int64(n))
    down.post(tap: .cghidEventTap)
    usleep(20000)
    let up = CGEvent(mouseEventSource: src, mouseType: .leftMouseUp, mouseCursorPosition: p, mouseButton: .left)!
    up.setIntegerValueField(.mouseEventClickState, value: Int64(n))
    up.post(tap: .cghidEventTap)
    usleep(40000)
  }
}

// MARK: - Pasteboard

func snapshotPasteboard() -> [[NSPasteboard.PasteboardType: Data]] {
  (NSPasteboard.general.pasteboardItems ?? []).map { item in
    var d: [NSPasteboard.PasteboardType: Data] = [:]
    for t in item.types { if let data = item.data(forType: t) { d[t] = data } }
    return d
  }
}

func restorePasteboard(_ snap: [[NSPasteboard.PasteboardType: Data]]) {
  let pb = NSPasteboard.general
  pb.clearContents()
  let items: [NSPasteboardItem] = snap.map { d in
    let it = NSPasteboardItem()
    for (t, data) in d { it.setData(data, forType: t) }
    return it
  }
  if !items.isEmpty { pb.writeObjects(items) }
}

// MARK: - Focus

/// Process pending run-loop work so NSWorkspace / NSRunningApplication state is current.
func refreshAppState() { RunLoop.current.run(mode: .default, before: Date()) }

/// Live check (via AX, not cached NSRunningApplication state) that Outlook is the frontmost app.
func outlookIsFrontmost() -> Bool {
  guard let app = try? appElement(), let v = attr(app, kAXFrontmostAttribute) else { return false }
  return CFGetTypeID(v) == CFBooleanGetTypeID() && (v as! Bool)
}

func requireFrontmost(_ what: String) throws {
  if !outlookIsFrontmost() { throw HelperError("Outlook is not the frontmost app; refusing to send \(what) to another application") }
}

func frontmostInfo() -> [String: Any] {
  guard let f = NSWorkspace.shared.frontmostApplication else { return [:] }
  return ["pid": f.processIdentifier, "bundleId": f.bundleIdentifier ?? "", "name": f.localizedName ?? ""]
}

func waitUntil(_ timeoutMs: Int, _ cond: () -> Bool) -> Bool {
  let deadline = Date().addingTimeInterval(Double(timeoutMs) / 1000)
  while Date() < deadline {
    if cond() { return true }
    RunLoop.current.run(until: Date().addingTimeInterval(0.04))
  }
  return cond()
}

// MARK: - Menus

func menuItem(_ menubar: AXUIElement, _ path: [String]) throws -> AXUIElement {
  var cur = menubar
  for (i, title) in path.enumerated() {
    var items = children(cur)
    if let menu = items.first(where: { str(attr($0, kAXRoleAttribute)) == "AXMenu" }) { items = children(menu) }
    // "..." and "…" are interchangeable in menu titles.
    let norm = { (s: String) in s.replacingOccurrences(of: "…", with: "...") }
    guard let next = items.first(where: { norm(str(attr($0, kAXTitleAttribute)) ?? "") == norm(title) }) else {
      throw HelperError("menu item not found: \(path[0...i].joined(separator: " > "))")
    }
    cur = next
  }
  return cur
}

// MARK: - Commands

func arg<T>(_ req: [String: Any], _ k: String) throws -> T {
  guard let v = req[k] as? T else { throw HelperError("missing or invalid argument '\(k)'") }
  return v
}

func handle(_ req: [String: Any]) throws -> Any {
  let cmd: String = try arg(req, "cmd")
  extraAttrs = (req["extra"] as? [String]) ?? []
  refreshAppState()
  switch cmd {
  case "ping":
    let trusted = AXIsProcessTrusted()
    let running = (try? outlookApp()) != nil
    return ["trusted": trusted, "outlookRunning": running, "pid": (try? outlookApp().processIdentifier) ?? 0, "version": (try? outlookApp().bundleURL.flatMap { Bundle(url: $0)?.infoDictionary?["CFBundleShortVersionString"] }) as Any? ?? NSNull()]

  case "release":
    handles.removeAll(keepingCapacity: true)
    return true

  case "windows":
    // AXWindows only lists windows on the active Space; AXMainWindow/AXFocusedWindow/AXChildren also reach
    // windows on other Spaces (e.g. a full-screen Outlook).
    let app = try appElement()
    var all: [AXUIElement] = []
    func add(_ e: AXUIElement?) { if let e = e, !all.contains(where: { CFEqual($0, e) }) { all.append(e) } }
    for w in (attr(app, kAXWindowsAttribute) as? [AXUIElement]) ?? [] { add(w) }
    for c in children(app) where str(attr(c, kAXRoleAttribute)) == "AXWindow" { add(c) }
    add(attr(app, kAXMainWindowAttribute).map { $0 as! AXUIElement })
    add(attr(app, kAXFocusedWindowAttribute).map { $0 as! AXUIElement })
    let main = attr(app, kAXMainWindowAttribute).map { $0 as! AXUIElement }
    return all.map { w -> [String: Any] in
      var s = summary(w, withFrame: true)
      if let m = main, CFEqual(m, w) { s["main"] = true }
      if let mini = attr(w, kAXMinimizedAttribute) as? Bool, mini { s["minimized"] = true }
      return s
    }

  case "get":
    let e = try resolve(try arg(req, "h"))
    var s = summary(e, withFrame: (req["frame"] as? Bool) ?? false)
    if let names = req["attrs"] as? [String] {
      var extra: [String: Any] = [:]
      for n in names { extra[n] = jsonValue(attr(e, n)) }
      s["attrs"] = extra
    }
    if (req["actions"] as? Bool) ?? false {
      var acts: CFArray?
      AXUIElementCopyActionNames(e, &acts)
      s["actions"] = (acts as? [String]) ?? []
    }
    if (req["attributeNames"] as? Bool) ?? false {
      var names: CFArray?
      AXUIElementCopyAttributeNames(e, &names)
      s["attributeNames"] = (names as? [String]) ?? []
    }
    return s

  case "children":
    let e = try resolve(try arg(req, "h"))
    let withFrame = (req["frame"] as? Bool) ?? false
    return children(e).map { summary($0, withFrame: withFrame) }

  case "elements":
    // Elements held in an array attribute (e.g. a table's AXRows, which includes rows not built as children yet).
    let e = try resolve(try arg(req, "h"))
    let name: String = try arg(req, "attr")
    let els = (attr(e, name) as? [AXUIElement]) ?? []
    let sub = (req["subtree"] as? Int) ?? 0
    let limit = (req["limit"] as? Int) ?? 5000
    return els.prefix(limit).map { tree($0, depth: 0, maxDepth: sub, maxChildren: 50, withFrame: false) }

  case "parent":
    let e = try resolve(try arg(req, "h"))
    guard let p = attr(e, kAXParentAttribute) else { return NSNull() }
    return summary(p as! AXUIElement)

  case "tree":
    let e = try resolve(try arg(req, "h"))
    return tree(e, depth: 0, maxDepth: (req["depth"] as? Int) ?? 4, maxChildren: (req["maxChildren"] as? Int) ?? 100,
                withFrame: (req["frame"] as? Bool) ?? false)

  case "find":
    let e = try resolve((req["h"] as? String) ?? "main")
    let m = Matcher(spec: (req["match"] as? [String: Any]) ?? [:])
    let res = find(e, m, maxDepth: (req["depth"] as? Int) ?? 25, limit: (req["limit"] as? Int) ?? 50)
    let withFrame = (req["frame"] as? Bool) ?? false
    if let sub = req["subtree"] as? Int {
      return res.map { tree($0, depth: 0, maxDepth: sub, maxChildren: (req["maxChildren"] as? Int) ?? 100, withFrame: withFrame) }
    }
    return res.map { summary($0, withFrame: withFrame) }

  case "waitFor":
    let root: String = (req["h"] as? String) ?? "main"
    let m = Matcher(spec: (req["match"] as? [String: Any]) ?? [:])
    let depth = (req["depth"] as? Int) ?? 25
    let gone = (req["gone"] as? Bool) ?? false
    var found: AXUIElement?
    let ok = waitUntil((req["timeoutMs"] as? Int) ?? 5000) {
      guard let r = try? resolve(root) else { return gone }
      found = find(r, m, maxDepth: depth, limit: 1).first
      return gone ? found == nil : found != nil
    }
    if gone { return ["ok": ok] }
    return ["ok": ok, "node": found.map { summary($0, withFrame: (req["frame"] as? Bool) ?? false) } as Any? ?? NSNull()]

  case "press":
    let e = try resolve(try arg(req, "h"))
    let action = (req["action"] as? String) ?? kAXPressAction
    var name = action
    if !action.hasPrefix("AX") { // custom action by display name, e.g. "Open event"
      var acts: CFArray?
      AXUIElementCopyActionNames(e, &acts)
      guard let full = ((acts as? [String]) ?? []).first(where: { $0 == action || $0.hasPrefix("Name:\(action)\n") }) else {
        throw HelperError("element has no action '\(action)'")
      }
      name = full
    }
    let err = AXUIElementPerformAction(e, name as CFString)
    // Outlook's SwiftUI controls often return an error code (-25200/-25204/-25205) yet perform the action,
    // so callers verify effects instead of trusting the code.
    return ["axError": err.rawValue]

  case "set":
    let e = try resolve(try arg(req, "h"))
    let name: String = try arg(req, "attr")
    let type = (req["type"] as? String) ?? "string"
    let value: CFTypeRef
    switch type {
    case "bool": value = ((req["value"] as? Bool) ?? false) ? kCFBooleanTrue : kCFBooleanFalse
    case "number": value = NSNumber(value: (req["value"] as? Double) ?? 0)
    case "date":
      guard let s = req["value"] as? String, let d = iso.date(from: s) else { throw HelperError("date value must be ISO-8601 with offset") }
      value = d as CFDate
    default: value = ((req["value"] as? String) ?? "") as CFString
    }
    let err = AXUIElementSetAttributeValue(e, name as CFString, value)
    return ["axError": err.rawValue]

  case "text":
    let e = try resolve(try arg(req, "h"))
    var links: [[String: String]] = []
    let role = str(attr(e, kAXRoleAttribute)) ?? ""
    var text: String? = nil
    if role == "AXWebArea", (req["markers"] as? Bool) ?? true { text = webAreaText(e) }
    let dfs = collectText(e, links: &links)
    if text == nil || text!.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty { text = dfs }
    return ["text": text ?? "", "links": links]

  case "menu":
    // path: ["File", "Save As..."]; walks the menu bar by exact titles.
    let path: [String] = try arg(req, "path")
    var cur = try menuItem(try resolve("menubar"), path)
    if let en = attr(cur, kAXEnabledAttribute) as? Bool, !en, !outlookIsFrontmost() {
      // Menus are disabled while Outlook is in the background (e.g. the user switched apps): reactivate once.
      try outlookApp().activate(options: [.activateAllWindows])
      _ = waitUntil(2000, { outlookIsFrontmost() })
      _ = waitUntil(300, { false })
      cur = try menuItem(try resolve("menubar"), path)
    }
    if let en = attr(cur, kAXEnabledAttribute) as? Bool, !en {
      throw HelperError("menu item disabled: \(path.joined(separator: " > "))\(outlookIsFrontmost() ? "" : " (Outlook is not the frontmost app)")")
    }
    if (req["check"] as? Bool) ?? false { return ["enabled": true] }
    return ["axError": AXUIElementPerformAction(cur, kAXPressAction as CFString).rawValue]

  case "frontmost":
    return frontmostInfo()

  case "activate":
    let prev = frontmostInfo()
    let app = try outlookApp()
    if !outlookIsFrontmost() {
      if app.isHidden { app.unhide() }
      app.activate(options: [.activateAllWindows])
      if !waitUntil(1500, { outlookIsFrontmost() }) {
        AXUIElementSetAttributeValue(try appElement(), kAXFrontmostAttribute as CFString, kCFBooleanTrue)
        _ = waitUntil(1500, { outlookIsFrontmost() })
      }
    }
    // Let the Space switch settle so menus validate and key events land.
    _ = waitUntil((req["settleMs"] as? Int) ?? 250, { false })
    return ["previous": prev, "active": outlookIsFrontmost()]

  case "restore":
    guard let pid = req["pid"] as? Int32 ?? (req["pid"] as? Int).map({ Int32($0) }),
          let app = NSRunningApplication(processIdentifier: pid) else { return ["restored": false] }
    if app.bundleIdentifier == outlookBundleId { return ["restored": true] }
    app.activate(options: [])
    let ok = waitUntil(1500, { app.isActive })
    if !ok { AXUIElementSetAttributeValue(AXUIElementCreateApplication(pid), kAXFrontmostAttribute as CFString, kCFBooleanTrue) }
    return ["restored": ok]

  case "key":
    // Only layout-independent keys (arrows, Return, Tab, Escape, Delete…) should be sent this way;
    // letter shortcuts depend on the keyboard layout (e.g. AZERTY) — use menus instead.
    let code = CGKeyCode((req["code"] as? Int) ?? 0)
    let mods = (req["mods"] as? [String]) ?? []
    let app = try outlookApp()
    try requireFrontmost("a key press")
    postKey(code, modifierFlags(mods), pid: app.processIdentifier, global: false)
    return true

  case "type":
    let text: String = try arg(req, "text")
    try typeText(text, pid: try outlookApp().processIdentifier)
    return true

  case "paste":
    // Put text (and optional HTML) on the pasteboard, invoke Edit > Paste (layout independent), then
    // restore the user's clipboard unless they copied something new meanwhile.
    let text: String = try arg(req, "text")
    let app = try appElement()
    try requireFrontmost("a paste")
    guard let mb = attr(app, kAXMenuBarAttribute) else { throw HelperError("no menu bar") }
    let pasteItem = try menuItem(mb as! AXUIElement, ["Edit", "Paste"])
    let snap = snapshotPasteboard()
    let pb = NSPasteboard.general
    pb.clearContents()
    let item = NSPasteboardItem()
    item.setString(text, forType: .string)
    if let html = req["html"] as? String { item.setString(html, forType: .html) }
    // Ask clipboard managers not to record this temporary content.
    item.setString("", forType: NSPasteboard.PasteboardType("org.nspasteboard.TransientType"))
    item.setString("", forType: NSPasteboard.PasteboardType("org.nspasteboard.ConcealedType"))
    pb.writeObjects([item])
    let ours = pb.changeCount
    defer {
      _ = waitUntil((req["settleMs"] as? Int) ?? 1200, { false })
      if NSPasteboard.general.changeCount == ours { restorePasteboard(snap) }
    }
    let err = AXUIElementPerformAction(pasteItem, kAXPressAction as CFString)
    return ["axError": err.rawValue]

  case "click":
    let e = try resolve(try arg(req, "h"))
    try requireFrontmost("a mouse click")
    _ = AXUIElementPerformAction(e, "AXScrollToVisible" as CFString)
    guard let f = frame(e) else { throw HelperError("element has no frame") }
    let p = CGPoint(x: f.midX + ((req["dx"] as? Double) ?? 0), y: f.midY + ((req["dy"] as? Double) ?? 0))
    // Make sure the point actually hits the target (or one of its descendants) inside Outlook.
    var hit: AXUIElement?
    guard AXUIElementCopyElementAtPosition(try appElement(), Float(p.x), Float(p.y), &hit) == .success, var cur = hit else {
      throw HelperError("click target is not visible")
    }
    var inside = false
    // SwiftUI may vend a different AXUIElement for the same node, so also compare role, label and frame.
    let sig = { (x: AXUIElement) -> String in
      "\(str(attr(x, kAXRoleAttribute)) ?? "")|\(str(attr(x, kAXDescriptionAttribute)) ?? "")|\(str(attr(x, kAXTitleAttribute)) ?? "")|\(frame(x).map { "\($0)" } ?? "")"
    }
    let target = sig(e)
    for _ in 0..<40 {
      if CFEqual(cur, e) || sig(cur) == target { inside = true; break }
      guard let parent = attr(cur, kAXParentAttribute) else { break }
      cur = parent as! AXUIElement
    }
    if !inside {
      throw HelperError("click target is covered by another element (hit \(str(attr(hit!, kAXRoleAttribute)) ?? "?") '\(str(attr(hit!, kAXDescriptionAttribute)) ?? str(attr(hit!, kAXTitleAttribute)) ?? "")')")
    }
    click(at: p, double: (req["double"] as? Bool) ?? false)
    return ["x": p.x, "y": p.y]

  case "sleep":
    _ = waitUntil((req["ms"] as? Int) ?? 100, { false })
    return true

  default:
    throw HelperError("unknown command '\(cmd)'")
  }
}

// MARK: - Main loop

func reply(_ obj: [String: Any]) {
  guard let data = try? JSONSerialization.data(withJSONObject: obj, options: []) else {
    print("{\"ok\":false,\"error\":\"unserializable response\"}")
    return
  }
  FileHandle.standardOutput.write(data)
  FileHandle.standardOutput.write("\n".data(using: .utf8)!)
}

let helperVersion = "0.1.0"
if CommandLine.arguments.contains("--version") {
  print("axlook-helper \(helperVersion)")
  exit(0)
}

if CommandLine.arguments.contains("--check") {
  let ok = AXIsProcessTrusted()
  print(ok ? "accessibility: trusted" : "accessibility: NOT trusted")
  exit(ok ? 0 : 1)
}

while let line = readLine(strippingNewline: true) {
  if line.isEmpty { continue }
  guard let data = line.data(using: .utf8),
        let req = (try? JSONSerialization.jsonObject(with: data)) as? [String: Any] else {
    reply(["ok": false, "error": "invalid JSON request"])
    continue
  }
  let id = req["id"] ?? NSNull()
  do {
    let result = try handle(req)
    reply(["id": id, "ok": true, "result": result])
  } catch let e as HelperError {
    reply(["id": id, "ok": false, "error": e.message])
  } catch {
    reply(["id": id, "ok": false, "error": String(describing: error)])
  }
}
