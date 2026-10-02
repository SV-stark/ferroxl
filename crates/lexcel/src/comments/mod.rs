//! Cell comments (`openpyxl/comments/`).

/// A comment attached to a cell.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Comment {
    text: String,
    author: String,
}

impl Comment {
    /// Build a comment from its text and author.
    pub fn new(text: impl Into<String>, author: impl Into<String>) -> Self {
        Comment {
            text: text.into(),
            author: author.into(),
        }
    }

    /// The name recorded for the author.
    pub fn author(&self) -> &str {
        &self.author
    }

    /// Replace the author.
    pub fn set_author(&mut self, author: impl Into<String>) {
        self.author = author.into();
    }

    /// The text of the comment.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Replace the comment text.
    pub fn set_text(&mut self, text: impl Into<String>) {
        self.text = text.into();
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accessors_round_trip() {
        let mut comment = Comment::new("hello", "eric");
        assert_eq!(comment.text(), "hello");
        assert_eq!(comment.author(), "eric");
        comment.set_text("bye");
        comment.set_author("gazoni");
        assert_eq!(comment.text(), "bye");
        assert_eq!(comment.author(), "gazoni");
    }

    #[test]
    fn equality_is_by_text_and_author() {
        assert_eq!(Comment::new("a", "b"), Comment::new("a", "b"));
        assert_ne!(Comment::new("a", "b"), Comment::new("a", "c"));
    }
}
