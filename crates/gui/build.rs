fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").unwrap_or_default() == "windows" {
        let icon_path = format!(
            "{}/assets/icon.ico",
            std::env::var("CARGO_MANIFEST_DIR").unwrap()
        );
        let mut res = winresource::WindowsResource::new();
        res.set_icon(&icon_path);
        if let Err(e) = res.compile() {
            eprintln!("warning: failed to set icon: {}", e);
        }
    }
}
