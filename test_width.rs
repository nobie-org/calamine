fn pixels_to_character_units(pixels: u32, mdw: f64, padding: u32) -> f64 {
    let noc = (pixels - padding) as f64 / mdw;
    let val = (noc * mdw + padding as f64) + ((noc * mdw + padding as f64) % 8.0);
    let result = (val - padding as f64) / mdw;
    (result * 256.0).trunc() / 256.0
}

fn main() {
    let result = pixels_to_character_units(61, 7.0, 5);
    println\!("Result: {}", result);
    println\!("Difference from 8.0: {}", (result - 8.0).abs());
}
