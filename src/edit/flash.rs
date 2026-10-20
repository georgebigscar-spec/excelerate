//! Flash Fill: a column filled by example.
//!
//! Given a few rows typed by hand next to some data, Flash Fill finds how
//! each was made from the cells to its left (which words, in what case, cut
//! how, with what between them) and makes the rest of the column the same
//! way. Ported from the Flash Fill of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0), and taught
//! to read several source columns, cut inside a word and step over a header.

use crate::coordinate::{CellRef, Col, Row};
use crate::error::{Error, Result};
use crate::model::{CellValue, Spreadsheet};

/// How many columns to the left Flash Fill reads at most.
const MAX_SOURCES: u32 = 16;

/// A case change.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Case {
    Same,
    Upper,
    Lower,
    Proper,
}

impl Case {
    const ALL: [Self; 4] = [Self::Same, Self::Upper, Self::Lower, Self::Proper];

    fn apply(self, text: &str) -> String {
        match self {
            Self::Same => text.to_owned(),
            Self::Upper => text.to_uppercase(),
            Self::Lower => text.to_lowercase(),
            Self::Proper => {
                let mut chars = text.chars();
                chars.next().map_or_else(String::new, |first| {
                    first
                        .to_uppercase()
                        .chain(chars.flat_map(char::to_lowercase))
                        .collect()
                })
            }
        }
    }
}

/// Which word of which source: `index` counts from the start, or from the
/// end when negative (-1 is the last).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Word {
    source: usize,
    index: i32,
}

/// One piece of an output.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Piece {
    /// Text that is the same in every row.
    Literal(String),
    /// A whole word, in some case.
    Word(Word, Case),
    /// The first letter of a word, in some case.
    Initial(Word, Case),
    /// Characters `from..to` of a word, counted from its start.
    Slice(Word, usize, usize),
}

/// A way of making an output from the cells left of it, learnt from
/// examples.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FlashFill(Vec<Piece>);

/// The words of a text: runs of letters, digits and apostrophes.
fn words(text: &str) -> Vec<String> {
    text.split(|c: char| !c.is_alphanumeric() && c != '\'')
        .filter(|w| !w.is_empty())
        .map(str::to_owned)
        .collect()
}

fn word_at(words: &[Vec<String>], at: Word) -> Option<&str> {
    let list = words.get(at.source)?;
    let i = if at.index >= 0 {
        usize::try_from(at.index).ok()?
    } else {
        list.len()
            .checked_sub(usize::try_from(at.index.unsigned_abs()).ok()?)?
    };
    list.get(i).map(String::as_str)
}

impl FlashFill {
    /// Learns how the outputs were made from their sources: each example is
    /// the texts of a row's source cells, left to right, and the text typed
    /// beside them. `None` when no single way explains every example.
    ///
    /// ```
    /// use excelerate::edit::FlashFill;
    ///
    /// let fill = FlashFill::learn(&[(vec!["Ada Lovelace"], "Lovelace, A.")]).unwrap();
    /// assert_eq!(fill.apply(&["Alan Turing"]).as_deref(), Some("Turing, A."));
    /// ```
    #[must_use]
    pub fn learn<S: AsRef<str>>(examples: &[(Vec<S>, &str)]) -> Option<Self> {
        let examples: Vec<(Vec<String>, String)> = examples
            .iter()
            .map(|(sources, out)| {
                (
                    sources.iter().map(|s| s.as_ref().to_owned()).collect(),
                    (*out).to_owned(),
                )
            })
            .collect();
        let (sources, output) = examples.first()?;
        candidates(sources, output).into_iter().find(|program| {
            examples
                .iter()
                .all(|(s, out)| program.apply(s).as_deref() == Some(out.as_str()))
        })
    }

    /// The output for one row's sources, or `None` when the row lacks a word
    /// the pattern needs.
    #[must_use]
    pub fn apply<S: AsRef<str>>(&self, sources: &[S]) -> Option<String> {
        let words: Vec<Vec<String>> = sources.iter().map(|s| words(s.as_ref())).collect();
        let mut out = String::new();
        for piece in &self.0 {
            match piece {
                Piece::Literal(text) => out.push_str(text),
                Piece::Word(at, case) => out.push_str(&case.apply(word_at(&words, *at)?)),
                Piece::Initial(at, case) => {
                    let first = word_at(&words, *at)?.chars().next()?;
                    out.push_str(&case.apply(&first.to_string()));
                }
                Piece::Slice(at, from, to) => {
                    let word: Vec<char> = word_at(&words, *at)?.chars().collect();
                    out.extend(word.get(*from..*to)?);
                }
            }
        }
        Some(out)
    }
}

/// The ways one example could have been made, best guesses first: the
/// output is cut, left to right, into the longest piece of a source word
/// that starts it, or a literal character.
fn candidates(sources: &[String], output: &str) -> Vec<FlashFill> {
    let words: Vec<Vec<String>> = sources.iter().map(|s| words(s)).collect();
    let mut out = Vec::new();
    for from_end in [false, true] {
        for slices in [false, true] {
            out.push(FlashFill(segment(&words, output, from_end, slices)));
        }
    }
    out.dedup();
    out
}

fn segment(words: &[Vec<String>], output: &str, from_end: bool, slices: bool) -> Vec<Piece> {
    let mut pieces: Vec<Piece> = Vec::new();
    let mut rest = output;
    while !rest.is_empty() {
        let mut best: Option<(usize, Piece)> = None;
        let offer = |best: &mut Option<(usize, Piece)>, len: usize, piece: Piece| {
            if len > 0 && best.as_ref().is_none_or(|(l, _)| len > *l) {
                *best = Some((len, piece));
            }
        };
        let place = |source: usize, i: i32, count: i32| Word {
            source,
            index: if from_end { i - count } else { i },
        };
        for (source, list) in words.iter().enumerate() {
            let count = i32::try_from(list.len()).unwrap_or(i32::MAX);
            for (i, word) in (0i32..).zip(list) {
                let at = place(source, i, count);
                for case in Case::ALL {
                    let cased = case.apply(word);
                    if rest.starts_with(&cased) {
                        offer(&mut best, cased.len(), Piece::Word(at, case));
                    }
                }
                if slices {
                    // The longest run of the word, two characters or more,
                    // that starts what is left.
                    let chars: Vec<char> = word.chars().collect();
                    for from in 0..chars.len() {
                        let run = chars[from..]
                            .iter()
                            .zip(rest.chars())
                            .take_while(|(a, b)| **a == *b)
                            .count();
                        if run >= 2 {
                            let len: usize =
                                chars[from..from + run].iter().map(|c| c.len_utf8()).sum();
                            offer(&mut best, len, Piece::Slice(at, from, from + run));
                        }
                    }
                }
            }
        }
        // An initial, only where no longer piece fits: "J. Smith".
        if best.as_ref().is_none_or(|(l, _)| *l <= 1) {
            'initials: for (source, list) in words.iter().enumerate() {
                let count = i32::try_from(list.len()).unwrap_or(i32::MAX);
                for (i, word) in (0i32..).zip(list) {
                    let Some(first) = word.chars().next() else {
                        continue;
                    };
                    for case in [Case::Upper, Case::Lower] {
                        let initial = case.apply(&first.to_string());
                        if rest.starts_with(&initial) {
                            best = Some((
                                initial.len(),
                                Piece::Initial(place(source, i, count), case),
                            ));
                            break 'initials;
                        }
                    }
                }
            }
        }
        if let Some((len, piece)) = best {
            pieces.push(piece);
            rest = &rest[len..];
        } else {
            let c = rest.chars().next().unwrap_or(' ');
            match pieces.last_mut() {
                Some(Piece::Literal(text)) => text.push(c),
                _ => pieces.push(Piece::Literal(c.to_string())),
            }
            rest = &rest[c.len_utf8()..];
        }
    }
    pieces
}

/// Fills the empty cells of `at`'s column by example: Excel's Data - Flash
/// Fill (Ctrl+E), run with `at` in the column to fill.
///
/// The sources are the columns just left of `at`, as many as hold data in
/// its row, up to 16. The rows are those around `at` where the sources hold
/// anything. The cells already filled in the column are the examples; when
/// no pattern explains them all, the first one is taken for a header and
/// left out. Each empty cell gets the pattern's output for its row as text,
/// as Excel writes it. Returns how many cells were filled: 0 when there is
/// no example, no column to read, or no pattern.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet.
pub fn flash_fill(book: &mut Spreadsheet, sheet: usize, at: CellRef) -> Result<usize> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let filled = |c: CellRef| {
        ws.get(c)
            .is_some_and(|cell| !cell.value.result().is_empty())
    };
    // The source columns: left of `at`, while they hold something in its row.
    let sources: Vec<Col> = (1..=MAX_SOURCES.min(at.col.index()))
        .map_while(|back| Col::new(at.col.index() - back))
        .take_while(|&col| filled(CellRef::new(col, at.row)))
        .collect::<Vec<_>>()
        .into_iter()
        .rev()
        .collect();
    if sources.is_empty() {
        return Ok(0);
    }
    let any_source = |row: u32| {
        Row::new(row).is_some_and(|r| sources.iter().any(|&c| filled(CellRef::new(c, r))))
    };
    let mut top = at.row.index();
    while top > 0 && any_source(top - 1) {
        top -= 1;
    }
    let mut bottom = at.row.index();
    while any_source(bottom + 1) {
        bottom += 1;
    }

    let mut examples: Vec<(Vec<String>, String)> = Vec::new();
    let mut todo: Vec<(CellRef, Vec<String>)> = Vec::new();
    for index in top..=bottom {
        let Some(row) = Row::new(index) else { continue };
        let texts: Vec<String> = sources
            .iter()
            .map(|&c| book.formatted(sheet, CellRef::new(c, row)))
            .collect();
        let target = CellRef::new(at.col, row);
        if filled(target) {
            examples.push((texts, book.formatted(sheet, target)));
        } else {
            todo.push((target, texts));
        }
    }
    let learn = |examples: &[(Vec<String>, String)]| {
        let borrowed: Vec<(Vec<&str>, &str)> = examples
            .iter()
            .map(|(s, o)| (s.iter().map(String::as_str).collect(), o.as_str()))
            .collect();
        FlashFill::learn(&borrowed)
    };
    let program = match learn(&examples) {
        Some(p) => p,
        // A header over the examples explains nothing.
        None if examples.len() > 1 => match learn(&examples[1..]) {
            Some(p) => p,
            None => return Ok(0),
        },
        None => return Ok(0),
    };

    let outputs: Vec<(CellRef, String)> = todo
        .into_iter()
        .filter_map(|(c, texts)| program.apply(&texts).map(|out| (c, out)))
        .collect();
    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    for (c, out) in &outputs {
        ws.entry(*c).value = CellValue::text(out.as_str());
    }
    Ok(outputs.len())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn learn(examples: &[(&[&str], &str)]) -> FlashFill {
        let owned: Vec<(Vec<&str>, &str)> =
            examples.iter().map(|(s, o)| (s.to_vec(), *o)).collect();
        FlashFill::learn(&owned).unwrap()
    }

    #[test]
    fn words_reordered_with_separators() {
        let p = learn(&[(&["John Smith"], "Smith, John")]);
        assert_eq!(p.apply(&["Ada Lovelace"]).as_deref(), Some("Lovelace, Ada"));
    }

    #[test]
    fn initials_and_case() {
        let p = learn(&[(&["jane doe"], "JD")]);
        assert_eq!(p.apply(&["alan turing"]).as_deref(), Some("AT"));
        let p = learn(&[(&["grace HOPPER"], "Grace Hopper")]);
        assert_eq!(
            p.apply(&["edsger DIJKSTRA"]).as_deref(),
            Some("Edsger Dijkstra")
        );
    }

    #[test]
    fn two_source_columns() {
        let p = learn(&[(&["Ada", "Lovelace"], "Lovelace A.")]);
        assert_eq!(p.apply(&["Alan", "Turing"]).as_deref(), Some("Turing A."));
    }

    #[test]
    fn slices_inside_a_word() {
        let p = learn(&[
            (&["5551234567"], "(555) 123-4567"),
            (&["2125550199"], "(212) 555-0199"),
        ]);
        assert_eq!(p.apply(&["4155550123"]).as_deref(), Some("(415) 555-0123"));
    }

    #[test]
    fn the_last_word_from_the_end() {
        let p = learn(&[(&["Mary Ann Evans"], "Evans"), (&["Ann Evans"], "Evans")]);
        assert_eq!(p.apply(&["George Eliot Jr"]).as_deref(), Some("Jr"));
    }

    #[test]
    fn nothing_explains_contradictions() {
        let owned = vec![(vec!["a b"], "a"), (vec!["c d"], "d")];
        assert_eq!(FlashFill::learn(&owned), None);
    }
}
