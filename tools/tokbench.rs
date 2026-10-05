// Tokenizer throughput benchmark: cargo run --release --example tokbench <file.xml>
use std::time::Instant;

#[path = "../src/xml.rs"]
#[allow(dead_code)]
mod xml;

fn main() {
    let path = std::env::args().nth(1).expect("xml file");
    let data = std::fs::read(&path).unwrap();
    let t = Instant::now();
    let mut r = xml::Reader::new(xml::SliceSource::new(&data));
    let (mut starts, mut texts) = (0usize, 0usize);
    loop {
        match r.next() {
            Ok(xml::Ev::Start) | Ok(xml::Ev::Empty) => starts += 1,
            Ok(xml::Ev::Text) => texts += 1,
            Ok(xml::Ev::End) => {}
            Ok(xml::Ev::Eof) => break,
            Err(e) => panic!("{:?}", e),
        }
    }
    let dt = t.elapsed().as_secs_f64();
    println!(
        "{} bytes, {} elements, {} texts in {:.3}s = {:.0} MB/s",
        data.len(),
        starts,
        texts,
        dt,
        data.len() as f64 / dt / 1e6
    );
}
