function toggleSidebar() {
    const sidebar = document.getElementById("sidebar");
    const main = document.getElementById("main");

    sidebar.classList.toggle("collapsed");
    main.classList.toggle("collapsed");
}
