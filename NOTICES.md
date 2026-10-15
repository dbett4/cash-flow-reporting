# Source and dependency notices

Copyright 2026 Dave Bettner. All rights reserved.

The reporting compiler and local Office builder were recovered from the owner's
existing private implementation at the commit recorded in
`SOURCE-PROVENANCE.json`. No license or copyright notice existed in the extracted
source files. No upstream open-source license grant is invented or implied.
Their schema identifiers retain their historical `acfr` prefix for compatibility;
the original private case data and native service adapters are excluded.

The Rust dependency graph is retained exactly in
`tools/report-compiler/Cargo.lock`. Direct dependencies are calamine,
rust_decimal, serde, serde_json and sha2. Python Office construction uses
openpyxl and python-docx. They and their transitive dependencies retain their
respective upstream licenses and notices in the installed distributions. This
repository does not vendor their implementation source or alter those licenses.

The fictional Harbor case, policies, expected amounts and public demonstration
were authored for this portfolio. They are not transformed client records.
