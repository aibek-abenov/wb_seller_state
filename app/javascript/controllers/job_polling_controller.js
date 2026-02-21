import { Controller } from "@hotwired/stimulus"

export default class extends Controller {
  static values = { url: String, redirect: String }

  connect() {
    this.poll()
  }

  disconnect() {
    clearTimeout(this.timer)
  }

  poll() {
    fetch(this.urlValue, {
      headers: { "Accept": "application/json" }
    })
      .then(response => response.json())
      .then(data => {
        if (data.status === "completed") {
          window.location.href = this.redirectValue
        } else if (data.status === "failed") {
          const el = document.getElementById("processing_status")
          if (el) {
            el.innerHTML =
              `<h2>Ошибка обработки</h2>` +
              `<p style="color:#ef4444;">${data.error || "Произошла ошибка. Попробуйте снова."}</p>` +
              `<a href="/dashboard" class="btn btn-primary">Вернуться на дашборд</a>`
          }
        } else {
          this.timer = setTimeout(() => this.poll(), 2000)
        }
      })
      .catch(() => {
        this.timer = setTimeout(() => this.poll(), 3000)
      })
  }
}
