use std::cell::RefCell;

use wasm_bindgen::prelude::*;

use sofdocs_core::document::editor::{
    self, DocPosition, DocSelection, EditOp, StyleChange, UndoStack,
};
use sofdocs_core::document::model::Document;
use sofdocs_core::document::parser::parse_docx as core_parse_docx;
use sofdocs_core::document::renderer::render_to_html;
use sofdocs_core::document::writer::write_docx as core_write_docx;

thread_local! {
    static DOC: RefCell<Document> = RefCell::new(Document::default());
    static UNDO: RefCell<UndoStack> = RefCell::new(UndoStack::new());
}

fn with_doc<F, R>(f: F) -> R
where
    F: FnOnce(&Document) -> R,
{
    DOC.with(|d| f(&d.borrow()))
}

fn with_doc_mut<F, R>(f: F) -> R
where
    F: FnOnce(&mut Document) -> R,
{
    DOC.with(|d| f(&mut d.borrow_mut()))
}

fn with_undo<F, R>(f: F) -> R
where
    F: FnOnce(&mut UndoStack) -> R,
{
    UNDO.with(|u| f(&mut u.borrow_mut()))
}

// --- Load / Parse ---

#[wasm_bindgen]
pub fn load_docx(bytes: &[u8]) -> Result<(), JsError> {
    let doc = core_parse_docx(bytes).map_err(|e| JsError::new(&e.to_string()))?;
    DOC.with(|d| *d.borrow_mut() = doc);
    UNDO.with(|u| *u.borrow_mut() = UndoStack::new());
    Ok(())
}

#[wasm_bindgen]
pub fn get_html() -> String {
    with_doc(render_to_html)
}

#[wasm_bindgen]
pub fn get_plain_text() -> String {
    with_doc(|d| d.to_plain_text())
}

#[wasm_bindgen]
pub fn get_word_count() -> usize {
    with_doc(|d| d.word_count())
}

#[wasm_bindgen]
pub fn get_paragraph_count() -> usize {
    with_doc(|d| d.paragraph_count())
}

#[wasm_bindgen]
pub fn get_document_json() -> Result<JsValue, JsError> {
    with_doc(|d| serde_wasm_bindgen::to_value(d).map_err(|e| JsError::new(&e.to_string())))
}

// --- Save ---

#[wasm_bindgen]
pub fn save_docx() -> Result<Vec<u8>, JsError> {
    with_doc(|d| core_write_docx(d).map_err(|e| JsError::new(&e.to_string())))
}

// --- Edit operations ---

#[wasm_bindgen]
pub fn insert_text(paragraph: usize, offset: usize, text: &str) -> String {
    let pos = DocPosition { paragraph, offset };
    let _end_pos = with_doc_mut(|d| editor::insert_text(d, pos, text));
    with_undo(|u| {
        u.push(EditOp::InsertText {
            position: pos,
            text: text.to_string(),
        })
    });
    get_html()
}

#[wasm_bindgen]
pub fn delete_range(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
) -> String {
    let sel = DocSelection {
        start: DocPosition {
            paragraph: start_para,
            offset: start_offset,
        },
        end: DocPosition {
            paragraph: end_para,
            offset: end_offset,
        },
    };
    let (_, deleted) = with_doc_mut(|d| editor::delete_range(d, sel));
    with_undo(|u| {
        u.push(EditOp::DeleteRange {
            selection: sel,
            deleted_content: deleted,
        })
    });
    get_html()
}

#[wasm_bindgen]
pub fn split_paragraph(paragraph: usize, offset: usize) -> String {
    let pos = DocPosition { paragraph, offset };
    with_doc_mut(|d| editor::split_paragraph(d, pos));
    with_undo(|u| u.push(EditOp::SplitParagraph { position: pos }));
    get_html()
}

#[wasm_bindgen]
pub fn toggle_bold(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
) -> String {
    apply_style_wasm(start_para, start_offset, end_para, end_offset, StyleChange::ToggleBold)
}

#[wasm_bindgen]
pub fn toggle_italic(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
) -> String {
    apply_style_wasm(start_para, start_offset, end_para, end_offset, StyleChange::ToggleItalic)
}

#[wasm_bindgen]
pub fn toggle_underline(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
) -> String {
    apply_style_wasm(start_para, start_offset, end_para, end_offset, StyleChange::ToggleUnderline)
}

#[wasm_bindgen]
pub fn toggle_strikethrough(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
) -> String {
    apply_style_wasm(
        start_para,
        start_offset,
        end_para,
        end_offset,
        StyleChange::ToggleStrikethrough,
    )
}

#[wasm_bindgen]
pub fn set_font_family(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
    font: &str,
) -> String {
    apply_style_wasm(
        start_para,
        start_offset,
        end_para,
        end_offset,
        StyleChange::SetFontFamily(font.to_string()),
    )
}

#[wasm_bindgen]
pub fn set_font_size(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
    size: f32,
) -> String {
    apply_style_wasm(
        start_para,
        start_offset,
        end_para,
        end_offset,
        StyleChange::SetFontSize(size),
    )
}

#[wasm_bindgen]
pub fn set_color(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
    color: &str,
) -> String {
    apply_style_wasm(
        start_para,
        start_offset,
        end_para,
        end_offset,
        StyleChange::SetColor(color.to_string()),
    )
}

#[wasm_bindgen]
pub fn set_alignment(paragraph: usize, alignment: &str) -> String {
    let align = if alignment.is_empty() {
        None
    } else {
        Some(alignment.to_string())
    };
    let old = with_doc_mut(|d| editor::set_alignment(d, paragraph, align.clone()));
    with_undo(|u| {
        u.push(EditOp::SetAlignment {
            paragraph,
            new_alignment: align,
            old_alignment: old,
        })
    });
    get_html()
}

// --- Undo / Redo ---

#[wasm_bindgen]
pub fn undo() -> String {
    DOC.with(|d| {
        UNDO.with(|u| {
            editor::undo(&mut d.borrow_mut(), &mut u.borrow_mut());
        })
    });
    get_html()
}

#[wasm_bindgen]
pub fn redo() -> String {
    DOC.with(|d| {
        UNDO.with(|u| {
            editor::redo(&mut d.borrow_mut(), &mut u.borrow_mut());
        })
    });
    get_html()
}

#[wasm_bindgen]
pub fn can_undo() -> bool {
    with_undo(|u| u.can_undo())
}

#[wasm_bindgen]
pub fn can_redo() -> bool {
    with_undo(|u| u.can_redo())
}

// --- Helpers ---

fn apply_style_wasm(
    start_para: usize,
    start_offset: usize,
    end_para: usize,
    end_offset: usize,
    change: StyleChange,
) -> String {
    let sel = DocSelection {
        start: DocPosition {
            paragraph: start_para,
            offset: start_offset,
        },
        end: DocPosition {
            paragraph: end_para,
            offset: end_offset,
        },
    };
    let previous = with_doc_mut(|d| editor::apply_style(d, sel, &change));
    with_undo(|u| {
        u.push(EditOp::ApplyStyle {
            selection: sel,
            style_change: change,
            previous_runs: previous,
        })
    });
    get_html()
}

// --- Header/Footer rendering ---

#[wasm_bindgen]
pub fn get_header_html() -> String {
    with_doc(|d| {
        if d.body.headers.is_empty() {
            return String::new();
        }
        let mut html = String::new();
        for header in &d.body.headers {
            for para in &header.paragraphs {
                html.push_str("<p>");
                for run in &para.runs {
                    html.push_str(&run.text);
                }
                html.push_str("</p>");
            }
        }
        html
    })
}

#[wasm_bindgen]
pub fn get_footer_html() -> String {
    with_doc(|d| {
        if d.body.footers.is_empty() {
            return String::new();
        }
        let mut html = String::new();
        for footer in &d.body.footers {
            for para in &footer.paragraphs {
                html.push_str("<p>");
                for run in &para.runs {
                    html.push_str(&run.text);
                }
                html.push_str("</p>");
            }
        }
        html
    })
}

#[wasm_bindgen]
pub fn get_page_count() -> usize {
    // Approximate page count based on paragraph count (rough heuristic)
    let para_count = with_doc(|d| d.paragraph_count());
    std::cmp::max(1, (para_count + 25) / 26)
}

// --- Legacy API (keep backward compat for existing callers) ---

#[wasm_bindgen]
pub fn parse_docx(bytes: &[u8]) -> Result<JsValue, JsError> {
    load_docx(bytes)?;
    get_document_json()
}

#[wasm_bindgen]
pub fn get_document_html(bytes: &[u8]) -> Result<String, JsError> {
    load_docx(bytes)?;
    Ok(get_html())
}

#[wasm_bindgen]
pub fn get_document_text(bytes: &[u8]) -> Result<String, JsError> {
    load_docx(bytes)?;
    Ok(get_plain_text())
}
