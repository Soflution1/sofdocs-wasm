use wasm_bindgen::prelude::*;

use sofdocs_core::document::parser::parse_docx as core_parse_docx;
use sofdocs_core::document::renderer::render_to_html;

/// Parse a .docx file from bytes and return the Document model as JSON.
#[wasm_bindgen]
pub fn parse_docx(bytes: &[u8]) -> Result<JsValue, JsError> {
    let doc = core_parse_docx(bytes).map_err(|e| JsError::new(&e.to_string()))?;
    serde_wasm_bindgen::to_value(&doc).map_err(|e| JsError::new(&e.to_string()))
}

/// Parse a .docx and return just the plain text content.
#[wasm_bindgen]
pub fn get_document_text(bytes: &[u8]) -> Result<String, JsError> {
    let doc = core_parse_docx(bytes).map_err(|e| JsError::new(&e.to_string()))?;
    Ok(doc.to_plain_text())
}

/// Parse a .docx and return an HTML representation for rendering.
#[wasm_bindgen]
pub fn get_document_html(bytes: &[u8]) -> Result<String, JsError> {
    let doc = core_parse_docx(bytes).map_err(|e| JsError::new(&e.to_string()))?;
    Ok(render_to_html(&doc))
}

/// Return the word count of a .docx file.
#[wasm_bindgen]
pub fn get_word_count(bytes: &[u8]) -> Result<usize, JsError> {
    let doc = core_parse_docx(bytes).map_err(|e| JsError::new(&e.to_string()))?;
    Ok(doc.word_count())
}
