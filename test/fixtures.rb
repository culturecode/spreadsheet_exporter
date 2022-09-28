def country_and_city
  country = sample_country
  city = CONDITIONAL_CITIES[country].sample
  {country: country, city: city}
end

def sample_country
  CONDITIONAL_CITIES.keys.sample
end

MEALS = %w[Omnivore Veg Vegan]

COUNTRIES = %w[Canada Turkey]

CONDITIONAL_CITIES = {
  COUNTRIES[0] => [
    "Sxwōxwiyám",
    "Toronto"
  ].map { |s| s.encode("UTF-8") },
  COUNTRIES[1] => [
    "Eskişehir",
    "İzmir",
    "İstanbul"
  ].map { |s| s.encode("UTF-8") }
}
