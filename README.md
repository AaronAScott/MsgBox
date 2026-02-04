# MsgBox.vb — A Drop‑In Replacement for VB’s MsgBox

A single‑file, dependency‑free replacement for the built‑in VB `MsgBox`.  
Same call syntax, same behavior, cleaner UI, optional theming.

**Note: this replacement purposefully *omits* support for the AbortRetryIgnore and RetryCancel styles.  If you need these fossils, you'll need to implement them yourself.**

To use, simply copy `MsgBox.vb` into your project and call `MsgBox` exactly as you always have.  
No rewrites. No configuration required.

## Optional Appearance Settings

If you want to customize the look, set any of the Shared properties in `MsgBoxTheme`:

- `MsgBoxTheme.ForeColor`
- `MsgBoxTheme.BackColor`
- `MsgBoxTheme.FontSize`  
  - Set to `0` to use the default small font (MS Sans Serif, 8.25pt) 
  - Any other value uses Arial at that size

If you don’t set anything, the module uses sensible defaults.

## What This Is

- One `.vb` file  
- No dependencies  
- No project file  
- No designer  
- No resources  
- No external configuration  
- No surprises  

Just drop it in and use it.

## License

Released under the “Do Anything You Want” License.  
See the LICENSE file for full text.