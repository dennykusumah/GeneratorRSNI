#         var s = pw.document.createElement('style');
#         s.id = '_mic_style';
#         s.textContent =
#           'div[data-testid="stChatInput"]{position:relative !important;}' +
#           'div[data-testid="stChatInput"] textarea{padding-right:128px !important;}' +
#           '#_mic_btn{' +
#             'position:absolute;right:52px;top:50%;transform:translateY(-50%);' +
#             'background:linear-gradient(135deg,#6366f1,#4f46e5);' +
#             'border:none;color:#fff;font-size:0.78rem;font-weight:700;' +
#             'padding:7px 14px;border-radius:99px;cursor:pointer;' +
#             'font-family:Outfit,sans-serif;' +
#             'box-shadow:0 3px 12px rgba(99,102,241,0.5);' +
#             'transition:all 0.2s;white-space:nowrap;' +
#             'display:inline-flex;align-items:center;gap:5px;z-index:200;' +
#           '}' +
#           '#_mic_btn:hover{background:linear-gradient(135deg,#818cf8,#6366f1);box-shadow:0 5px 18px rgba(99,102,241,0.65);}' +
#           '#_mic_btn.listening{background:linear-gradient(135deg,#ef4444,#dc2626)!important;animation:_mpulse 1s infinite;}' +
#           '@keyframes _mpulse{0%{box-shadow:0 0 0 0 rgba(239,68,68,0.6)}70%{box-shadow:0 0 0 8px rgba(239,68,68,0)}100%{box-shadow:0 0 0 0 rgba(239,68,68,0)}}' +
#           '#_mic_toast{display:none;position:fixed;bottom:72px;left:50%;transform:translateX(-50%);' +
#             'background:rgba(15,23,42,0.96);border:1px solid rgba(99,102,241,0.4);' +
#             'border-radius:12px;padding:7px 16px;font-size:0.78rem;color:#c7d2fe;' +
#             'font-family:sans-serif;z-index:99999;box-shadow:0 4px 20px rgba(0,0,0,0.4);' +
#             'white-space:nowrap;pointer-events:none;max-width:90vw;overflow:hidden;text-overflow:ellipsis;}';
#         pw.document.head.appendChild(s);
#       }

#       if (!pw.document.getElementById('_mic_toast')) {
#         var t = pw.document.createElement('div');
#         t.id = '_mic_toast';
#         pw.document.body.appendChild(t);
#       }
#       activeToast = pw.document.getElementById('_mic_toast');

#       function doInject() {
#         var chatInput = pw.document.querySelector('div[data-testid="stChatInput"]');
#         if (!chatInput) return false;
#         var existing = pw.document.getElementById('_mic_btn');
#         if (existing && chatInput.contains(existing)) {
#           activeBtn = existing;
#           existing.onclick = function(e){ e.preventDefault(); window.toggleMic(); };
#           fbBtn.style.display = 'none';
#           return true;
#         }
#         if (existing) existing.remove();
#         var btn = pw.document.createElement('button');
#         btn.id   = '_mic_btn';
#         btn.type = 'button';
#         btn.innerHTML = '🎤 Bicara';
#         btn.onclick = function(e){ e.preventDefault(); window.toggleMic(); };
#         chatInput.appendChild(btn);
#         activeBtn   = btn;
#         fbBtn.style.display = 'none';
#         return true;
#       }

#       if (!doInject()) {
#         var obs = new pw.MutationObserver(function(){ doInject(); });
#         obs.observe(pw.document.body, { childList:true, subtree:true });
#       } else {
#         var obs2 = new pw.MutationObserver(function(){ doInject(); });
#         obs2.observe(pw.document.body, { childList:true, subtree:true });
#       }

#     } catch(e) {
#       fbBtn.style.display = 'flex';
#     }
#   }

#   setupParentBtn();
# })();
# </script>
# """, height=56)

#     # Input teks
#     user_input = st.chat_input("Tanya seputar isi dokumen RSNI...")

#     if user_input:
#         st.session_state['_chat_history'].append({'role': 'user', 'content': user_input})
#         with st.chat_message('user', avatar='🧑'):
#             st.markdown(user_input)

#         with st.chat_message('assistant', avatar='🤖'):
#             with st.spinner("asistant sedang menganalisis..."):
#                 if has_doc:
#                     doc_ctx = _build_doc_context(sections, max_chars=14000)
#                     system_prompt = (
#                         "Kamu adalah asisten ahli RSNI yang membantu pengguna memahami dokumen. "
#                         "Jawab HANYA berdasarkan isi dokumen berikut. "
#                         "Gunakan Bahasa Indonesia yang jelas, terstruktur, dan akurat. "
#                         "Jika informasi tidak ada dalam dokumen, katakan dengan jujur.\n\n"
#                         f"=== ISI DOKUMEN ===\n{doc_ctx}\n==================="
#                     )
#                 else:
#                     system_prompt = (
#                         "Kamu adalah asisten ahli RSNI dan dokumen teknis BSN. "
#                         "Jawab dalam Bahasa Indonesia dengan jelas dan akurat. "
#                         "Belum ada dokumen — jawab berdasarkan pengetahuan umum SNI, ISO, IEC, dan standar lainnya."
#                     )
#                 api_messages = [
#                     {"role": m['role'], "content": m['content']}
#                     for m in st.session_state['_chat_history']
#                 ]
#                 reply = _claude_chat(system_prompt, api_messages)

#             st.markdown(reply)
#             _clean = re.sub(r'[*_`#>\-]+', '', reply)
#             _clean = re.sub(r'\s+', ' ', _clean).strip()
#             _components.html(_tts_html(_clean, f"new_{int(time.time())}"), height=44)
#             st.session_state['_chat_history'].append({'role': 'assistant', 'content': reply})

#     if st.session_state.get('_chat_history'):
#         if st.button("🗑️ Hapus Riwayat", key="clear_chat"):
#             st.session_state['_chat_history'] = []
#             st.rerun()

# Kondisi tanpa proses aktif atau halaman hasil: footer berada di akhir alur
# halaman. Saat proses aktif fungsi ini no-op karena sudah dirender di bawah
# progress/timer sebelum pekerjaan berat dimulai.
_render_footer_once()
