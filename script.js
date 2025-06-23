const searchText = document.getElementById("search-text");
const searchBtn = document.getElementById("searchBtn");
const searchResult = document.getElementById("search-result");
const searchMore = document.getElementById("show-more");

let newKeyword = "";
let page = 1;
const accessKey = "U9V8-lX9Gox1YByP3klXTGaDxaVW2UM1nqI_svG42-U";

async function imageSearch() {
  newKeyword = searchText.value;
  const url = `https://api.unsplash.com/search/photos?page=${page}&query=${newKeyword}&client_id=${accessKey}&per_page=12`;
  const response = await fetch(url);
  const data = await response.json();

  const results = data.results;

  results.map((result) => {
    const imagesRes = document.createElement("img");
    imagesRes.src = result.urls.small;
    searchResult.appendChild(imagesRes);
  });

  searchMore.style.visibility = "visible";

  console.log(data);
}

searchBtn.addEventListener("click", () => {
    searchResult.innerHTML = "";
  imageSearch();
});

searchMore.addEventListener("click", () => {
    page++;
    imageSearch();
})
