// concepts.h - concepts
#pragma once
#include <concepts>
#include <type_traits>

template<class B1, class B2, class D>
concept either_base_of_v = (std::is_base_of_v<B1, D> || std::is_base_of_v<B2, D>);

template<class B, class D1, class D2>
concept both_derived_from_v = (std::is_base_of_v<B, D1> && std::is_base_of_v<B, D2>);
