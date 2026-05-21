mod registry;
mod runtime;
mod safety;
mod schemas;

fn main() {
    if let Err(error) = runtime::run(
        std::env::args().collect(),
        std::env::current_dir().unwrap_or_default(),
    ) {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
