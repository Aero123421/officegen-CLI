mod registry;
mod runtime;
mod safety;
mod schemas;
mod v5_ooxml;
mod v5_workflow;
mod v5_xlsx_template;

fn main() {
    if let Err(error) = runtime::run(
        std::env::args().collect(),
        std::env::current_dir().unwrap_or_default(),
    ) {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
